using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using TagDiff.Core.Common;
using TagDiff.Core.Output;

namespace TagDiff.Core.Comparer
{
    public abstract class IgxlSheetComparerBase(string sheetName, string baseFile, string compFile, string firstRowKey, HashSet<string> skipColumns) : SheetComparerBase(sheetName, baseFile, compFile, firstRowKey, skipColumns)
    {
        protected string BaseSheetVersion { get; set; } = string.Empty;
        protected string CompareSheetVersion { get; set; } = string.Empty;

        protected List<string> KeyColHeaders { get; set; } = [];
        protected HashSet<string> UseList { get; set; } = new HashSet<string>(StringExtensions.IgnoreCase);

        public override List<CompareReport> Compare()
        {
            var result = new CompareReport();

            if (CompSheet.Count == 0)
            {
                result.SheetName = SheetName;
                result.SheetType = SheetType;
                result.TotalCount = BaseSheet.Count;
                result.ManualCount = BaseSheet.Count;
                result.SemiAutoCount = 0;
                return [result];
            }

            Dictionary<string, List<Location>> dicCompSheetItems = LoadSheetItems(CompSheet, CompSheetStartRow, CompSheetEndRow, CompSheetStartCol, CompSheetEndCol);
            int[] baseKeyIndices = BuildKeyColIndices(BaseSheet, KeyColHeaders, BaseSheetStartRow, BaseSheetStartCol, BaseSheetEndCol);

            int totalCount = 0;
            int failCount = 0;
            int semiAutoCount = 0;

            for (int row = BaseSheetStartRow + 1; row <= BaseSheetEndRow; row++)
            {
                string itemKey = GetItemKey(BaseSheet, row, baseKeyIndices).Trim();
                if (ByPassCheck(itemKey))
                {
                    continue;
                }

                totalCount++;
                if (dicCompSheetItems.TryGetValue(itemKey, out List<Location>? locations))
                {
                    Location? compLocation = locations.FirstOrDefault(l => !l.IsUsed);
                    if (compLocation == null)
                    {
                        failCount++;
                        SetCellSafe(BaseSheet, row, 0, "F");
                    }
                    else if (!CompareSheetRow(new Location(row, BaseSheetStartCol), compLocation, Skips))
                    {
                        SetCellSafe(BaseSheet, row, 0, "M");
                        semiAutoCount++;
                    }
                    else
                    {
                        SetCellSafe(BaseSheet, row, 0, "P");
                    }

                    if (compLocation != null)
                    {
                        compLocation.IsUsed = true;
                    }
                }
                else
                {
                    SetCellSafe(BaseSheet, row, 0, "F");
                    failCount++;
                }
            }

            IEnumerable<string> lines = BaseSheet.Select(row => string.Join("\t", row));
            File.WriteAllLines(BaseFile, lines);

            result.SheetName = SheetName;
            result.SheetType = SheetType;
            result.TotalCount = totalCount;
            result.ManualCount = failCount;
            result.SemiAutoCount = semiAutoCount;

            return [result];
        }

        protected bool ByPassCheck(string itemKey)
        {
            if (string.IsNullOrEmpty(itemKey) || (UseList.Count != 0 && !UseList.Contains(itemKey)))
            {
                return true;
            }

            return false;
        }

        protected static string NormalizeCell(string raw)
        {
            return (raw ?? string.Empty).Trim().ToUpperInvariant();
        }

        protected static int[] BuildKeyColIndices(List<string[]> sheet, List<string> keyColHeaders, int headerRow, int startCol, int endCol)
        {
            if (sheet == null || sheet.Count == 0 || keyColHeaders == null || keyColHeaders.Count == 0)
            {
                return [];
            }

            string[] header = sheet[headerRow];
            var indices = new List<int>(keyColHeaders.Count);
            for (int col = startCol; col <= endCol && col < header.Length; col++)
            {
                string h = NormalizeCell(header[col]);
                if (keyColHeaders.Exists(s => s.EqualsIgnoreCase(h)))
                {
                    indices.Add(col);
                }
            }
            return [.. indices];
        }

        protected static string GetItemKey(List<string[]> sheet, int row, int[] keyColIndices)
        {
            if (keyColIndices.Length == 0 || row >= sheet.Count)
            {
                return string.Empty;
            }

            string[] rowData = sheet[row];
            if (keyColIndices.Length == 1)
            {
                int col = keyColIndices[0];
                return col < rowData.Length ? NormalizeCell(rowData[col]) : string.Empty;
            }

            string[] keys = new string[keyColIndices.Length];
            for (int i = 0; i < keyColIndices.Length; i++)
            {
                int col = keyColIndices[i];
                keys[i] = col < rowData.Length ? NormalizeCell(rowData[col]) : string.Empty;
            }
            return string.Join(";", keys);
        }

        protected string GetItemKey(List<string[]> sheet, int row, int headerRow, int headerStartCol, int headerEndCol)
        {
            if (sheet == null || KeyColHeaders == null || KeyColHeaders.Count == 0)
            {
                return string.Empty;
            }

            int[] indices = BuildKeyColIndices(sheet, KeyColHeaders, headerRow, headerStartCol, headerEndCol);
            return GetItemKey(sheet, row, indices);
        }

        protected virtual Dictionary<string, List<Location>> LoadSheetItems(List<string[]> sheet, int sheetStartRow, int sheetEndRow, int sheetStartCol, int sheetEndCol)
        {
            var dicSheetItems = new Dictionary<string, List<Location>>(StringExtensions.IgnoreCase);
            if (sheet == null)
            {
                return dicSheetItems;
            }

            int[] keyIndices = BuildKeyColIndices(sheet, KeyColHeaders, sheetStartRow, sheetStartCol, sheetEndCol);
            for (int row = sheetStartRow + 1; row <= sheetEndRow; row++)
            {
                string itemKey = GetItemKey(sheet, row, keyIndices).Trim();
                if (!string.IsNullOrEmpty(itemKey))
                {
                    if (!dicSheetItems.TryGetValue(itemKey, out List<Location>? list))
                    {
                        list = [];
                        dicSheetItems[itemKey] = list;
                    }

                    list.Add(new Location(row, sheetStartCol));
                }
            }
            return dicSheetItems;
        }
    }
}
