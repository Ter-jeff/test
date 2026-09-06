using System;

using OfficeOpenXml;

using SkiaSharp;

namespace CommonLib.Extension
{
    /// <summary>
    /// GDI+-free column width estimation used by <see cref="EpplusExtensions"/> when the native
    /// EPPlus AutoFit (which needs GDI+/libgdiplus) is unavailable — e.g. macOS/Linux without libgdiplus.
    /// Primary path measures glyph advances with SkiaSharp (self-contained native binaries); if Skia's
    /// native load fails it degrades to a pure-managed proportional estimate. Never throws.
    /// </summary>
    internal static class EpplusColumnMeasurer
    {
        private static readonly object _sync = new();

        private static bool _skiaChecked;
        private static bool _skiaAvailable;
        private static SKFont? _font;
        private static SKFont? _fontBold;
        private static float _maxDigitWidthPx; // width of '0' — Excel's column-width unit

        // Excel's default workbook font is Calibri 11pt; convert pt -> px at 96 DPI.
        private const float DefaultFontSizePx = 11f * 96f / 72f;
        private const double MinWidth = 2.0;
        private const double MaxWidth = 255.0;   // Excel's hard column-width cap
        private const double PaddingUnits = 1.5; // small cushion so text isn't clipped
        private const int MaxRowsToScan = 5000;  // bound cost on very tall sheets

        /// <summary>Estimate and set the width of a single column over the worksheet's used row range.</summary>
        public static void EstimateColumn(ExcelWorksheet worksheet, int column)
        {
            if (worksheet?.Dimension == null || column < 1)
            {
                return;
            }

            try
            {
                int startRow = worksheet.Dimension.Start.Row;
                int endRow = Math.Min(worksheet.Dimension.End.Row, worksheet.Dimension.Start.Row + MaxRowsToScan);

                double best = 0;
                for (int row = startRow; row <= endRow; row++)
                {
                    ExcelRange cell = worksheet.Cells[row, column];
                    string text = cell.Text;
                    if (string.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    double units = MeasureWidthUnits(text, cell.Style.Font.Bold);
                    if (units > best)
                    {
                        best = units;
                    }
                }

                if (best <= 0)
                {
                    return;
                }

                worksheet.Column(column).Width = Math.Min(MaxWidth, Math.Max(MinWidth, best + PaddingUnits));
            }
            catch (Exception)
            {
                // Best-effort cosmetic sizing only; never let estimation break report generation.
            }
        }

        /// <summary>Width of <paramref name="text"/> in Excel column-width units (1 unit ≈ one '0' glyph).</summary>
        private static double MeasureWidthUnits(string text, bool bold)
        {
            EnsureSkia();
            if (_skiaAvailable)
            {
                lock (_sync)
                {
                    if (_skiaAvailable)
                    {
                        try
                        {
                            SKFont font = bold ? _fontBold! : _font!;
                            return font.MeasureText(text) / _maxDigitWidthPx;
                        }
                        catch (Exception)
                        {
                            _skiaAvailable = false; // fall through to managed estimate
                        }
                    }
                }
            }

            return BucketWidthUnits(text, bold);
        }

        private static void EnsureSkia()
        {
            if (_skiaChecked)
            {
                return;
            }

            lock (_sync)
            {
                if (_skiaChecked)
                {
                    return;
                }

                _skiaChecked = true;
                try
                {
                    // Calibri isn't present on macOS/Linux; Skia falls back to a default sans-serif,
                    // which is close enough for cosmetic column sizing.
                    _font = new SKFont(SKTypeface.FromFamilyName("Calibri"), DefaultFontSizePx);
                    _fontBold = new SKFont(SKTypeface.FromFamilyName("Calibri", SKFontStyle.Bold), DefaultFontSizePx);
                    _maxDigitWidthPx = _font.MeasureText("0");
                    if (_maxDigitWidthPx <= 0)
                    {
                        _maxDigitWidthPx = DefaultFontSizePx * 0.5f;
                    }

                    _skiaAvailable = true;
                }
                catch (Exception)
                {
                    _skiaAvailable = false;
                }
            }
        }

        /// <summary>
        /// Pure-managed fallback: sum per-glyph advance factors relative to a digit's width.
        /// Proportional (narrow 'i' vs wide 'W') but approximate; used only if SkiaSharp can't load.
        /// </summary>
        private static double BucketWidthUnits(string text, bool bold)
        {
            double units = 0;
            foreach (char c in text)
            {
                units += CharWidthFactor(c);
            }

            return bold ? units * 1.05 : units;
        }

        private static double CharWidthFactor(char c)
        {
            if (c == ' ')
            {
                return 0.5;
            }

            if ("iIl.,:;'|!ft()[]{}j".IndexOf(c) >= 0)
            {
                return 0.45;
            }

            if ("mwMW@%".IndexOf(c) >= 0)
            {
                return 1.5;
            }

            if (char.IsUpper(c))
            {
                return 1.15;
            }

            return 0.95;
        }
    }
}
