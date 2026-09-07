using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using CommonLib.Extension;

using IgxlLib.Enums;

using TagDiff.Core.Common;
using TagDiff.Core.Output;

using TagDiffCore.Utility;

namespace TagDiff.Core.Comparer
{
    public abstract class SheetComparerBase
    {
        protected const double Epsilon = 1e-12;

        public string SheetName { get; set; }
        public EnumSheetType SheetType { get; set; }

        protected string BaseFile { get; private set; }
        protected string CompFile { get; private set; }

        protected List<string[]> BaseSheet { get; set; }
        protected List<string[]> CompSheet { get; set; }

        protected int BaseSheetStartRow { get; set; }
        protected int BaseSheetEndRow { get; set; }
        protected int BaseSheetStartCol { get; set; }
        protected int BaseSheetEndCol { get; set; }

        protected int CompSheetStartRow { get; set; }
        protected int CompSheetEndRow { get; set; }
        protected int CompSheetStartCol { get; set; }
        protected int CompSheetEndCol { get; set; }
        public HashSet<int> Skips { get; set; }

        protected SheetComparerBase(string sheetName, string baseFile, string compFile, string firstRowKey, HashSet<string> skipColumns)
        {
            SheetName = sheetName;
            BaseFile = baseFile;
            CompFile = compFile;

            List<string[]> baseSheet = [];
            List<string[]> compareSheet = [];
            Parallel.Invoke(
                () => baseSheet = File.Exists(baseFile) ? [.. File.ReadLines(baseFile).Select(x => x.Split('\t'))] : [],
                () => compareSheet = File.Exists(compFile) ? [.. File.ReadLines(compFile).Select(x => x.Split('\t'))] : []
            );

            BaseSheet = baseSheet;
            CompSheet = compareSheet;

            if (baseSheet.Count > 0)
            {
                BaseSheetEndRow = baseSheet.Count - 1;
                BaseSheetEndCol = baseSheet[0].Length - 1;
            }

            if (compareSheet.Count > 0)
            {
                CompSheetEndRow = compareSheet.Count - 1;
                CompSheetEndCol = compareSheet[0].Length - 1;
            }

            SheetType = EnumSheetType.DTUnknown;

            GetDimension(firstRowKey);

            Skips = GetSkipColIndex(skipColumns);
        }

        public abstract List<CompareReport> Compare();

        protected virtual bool CompareSheetRow(Location baseLocation, Location compLocation, HashSet<int>? skips = null)
        {
            bool result = true;
            string[] baseRow = BaseSheet[baseLocation.Row];
            string[]? compRow = compLocation.Row >= 0 && compLocation.Row < CompSheet.Count ? CompSheet[compLocation.Row] : null;
            int compRowLen = compRow?.Length ?? 0;
            for (int col = baseLocation.Col; col <= BaseSheetEndCol; col++)
            {
                if (skips != null && skips.Contains(col))
                {
                    continue;
                }

                string baseContext = col < baseRow.Length ? baseRow[col] ?? string.Empty : string.Empty;
                string comContext = compRow != null && col < compRowLen ? compRow[col] ?? string.Empty : string.Empty;

                string baseTrim = baseContext.Trim().TrimStart('=');
                string comTrim = comContext.Trim().TrimStart('=');
                if (!baseTrim.EqualsIgnoreCase(comTrim))
                {
                    if (!(double.TryParse(baseTrim, NumberStyles.Any, CultureInfo.InvariantCulture, out double baseValue) && double.TryParse(comTrim, NumberStyles.Any, CultureInfo.InvariantCulture, out double comValue) && Math.Abs(baseValue - comValue) < Epsilon))
                    {
                        SetCellSafe(BaseSheet, baseLocation.Row, col, $"{baseContext} => {comContext}");
                        result = false;
                    }
                }
            }
            return result;
        }

        public static string GetCellSafe(List<string[]> sheet, int row, int col)
        {
            if (sheet == null || row < 0 || row >= sheet.Count)
            {
                return string.Empty;
            }

            string[] rowList = sheet[row];
            if (rowList == null || col < 0 || col >= rowList.Length)
            {
                return string.Empty;
            }

            return rowList[col] ?? string.Empty;
        }

        protected static void SetCellSafe(List<string[]> baseSheet, int rowIndex, int colIndex, string text)
        {
            ArgumentNullException.ThrowIfNull(baseSheet);
            if (rowIndex < 0 || rowIndex >= baseSheet.Count)
            {
                // avoid throwing for invalid indices during compare runs
                return;
            }

            if (colIndex < 0)
            {
                return;
            }

            string[] row = baseSheet[rowIndex];
            if (colIndex >= row.Length)
            {
                Array.Resize(ref row, colIndex + 1);
                baseSheet[rowIndex] = row;
            }
            baseSheet[rowIndex][colIndex] = text;
        }

        private void GetDimension(string firstRowKey)
        {
            (int startRow1, int endRow1, int startCol1, int endCol1) = SheetProvider.GetDimensions(BaseSheet, firstRowKey, TagDiffMain.IncludingBackup);
            BaseSheetStartRow = startRow1;
            BaseSheetEndRow = endRow1;
            BaseSheetStartCol = startCol1;
            BaseSheetEndCol = endCol1;

            (int startRow2, int endRow2, int startCol2, int endCol2) = SheetProvider.GetDimensions(CompSheet, firstRowKey, TagDiffMain.IncludingBackup);
            CompSheetStartRow = startRow2;
            CompSheetEndRow = endRow2;
            CompSheetStartCol = startCol2;
            CompSheetEndCol = endCol2;
        }

        protected HashSet<int> GetSkipColIndex(HashSet<string> skipColumns)
        {
            var skips = new HashSet<int>();

            if (BaseSheet == null || BaseSheet.Count == 0)
            {
                return skips;
            }

            string[] headerRow = BaseSheet[BaseSheetStartRow];

            for (int col = BaseSheetStartCol; col <= BaseSheetEndCol && col < headerRow.Length; col++)
            {
                string value = headerRow[col]?.Trim() ?? string.Empty;

                if (skipColumns.Contains(value))
                {
                    skips.Add(col);
                }
            }

            return skips;
        }
    }
}
