using System;
using System.Reflection;

using OfficeOpenXml;

namespace CommonLib.Extension
{
    public static partial class EpplusExtensions
    {
        // AutoFit requires GDI+ (libgdiplus on non-Windows). Probe once: if the graphics
        // stack is unavailable the first call throws (TypeInitializationException wrapping
        // DllNotFoundException, then TypeInitializationException on later calls), so we
        // remember it and switch to GDI+-free width estimation (EpplusColumnMeasurer) thereafter.
        private static bool _autoFitSupported = true;

        // ExcelColumn (EPPlus 4.5.3.2) doesn't publicly expose its owning worksheet; grab it via
        // the private field so the estimation fallback can read the column's cells. Cached; guarded.
        private static readonly FieldInfo? _columnWorksheetField =
            typeof(ExcelColumn).GetField("_worksheet", BindingFlags.Instance | BindingFlags.NonPublic);

        /// <summary>
        /// Actively probes once whether native (GDI+) AutoFit works, and caches the result.
        /// Call at startup to decide whether to warn the user. Returns false on macOS/Linux
        /// without libgdiplus; column sizing then falls back to estimation.
        /// </summary>
        public static bool IsNativeAutoFitAvailable()
        {
            if (!_autoFitSupported)
            {
                return false;
            }

            try
            {
                using var pkg = new ExcelPackage();
                ExcelWorksheet ws = pkg.Workbook.Worksheets.Add("probe");
                ws.Cells[1, 1].Value = "probe";
                ws.Column(1).AutoFit();
                return true;
            }
            catch (Exception)
            {
                _autoFitSupported = false;
                return false;
            }
        }

        /// <summary>Best-effort AutoFit for a single column; falls back to GDI+-free estimation. Never throws.</summary>
        public static bool TryAutoFit(this ExcelColumn column)
        {
            if (_autoFitSupported)
            {
                try
                {
                    column.AutoFit();
                    return true;
                }
                catch (Exception)
                {
                    _autoFitSupported = false;
                }
            }

            EstimateColumn(column);
            return false;
        }

        /// <summary>Best-effort AutoFitColumns for a range; falls back to GDI+-free estimation. Never throws.</summary>
        public static bool TryAutoFitColumns(this ExcelRange range)
        {
            if (_autoFitSupported)
            {
                try
                {
                    range.AutoFitColumns();
                    return true;
                }
                catch (Exception)
                {
                    _autoFitSupported = false;
                }
            }

            if (range != null)
            {
                for (int col = range.Start.Column; col <= range.End.Column; col++)
                {
                    EpplusColumnMeasurer.EstimateColumn(range.Worksheet, col);
                }
            }

            return false;
        }

        private static void EstimateColumn(ExcelColumn column)
        {
            if (column == null || _columnWorksheetField == null)
            {
                return;
            }

            try
            {
                if (_columnWorksheetField.GetValue(column) is ExcelWorksheet worksheet)
                {
                    EpplusColumnMeasurer.EstimateColumn(worksheet, column.ColumnMin);
                }
            }
            catch (Exception)
            {
                // Estimation is cosmetic; never break report generation over a column width.
            }
        }
    }
}
