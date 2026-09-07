using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;

using TagDiff.Core.Common;
using TagDiff.Core.Output;

using TagDiffCore.Regs;
using TagDiffCore.Utility;

namespace TagDiff.Core.Comparer
{
    public class AcSpecsSheetComparer : IgxlSheetComparerBase
    {
        private Dictionary<string, int> _baseDicCategoryColIndex = new(StringExtensions.IgnoreCase);
        private Dictionary<string, int> _compareDicCategoryColIndex = new(StringExtensions.IgnoreCase);

        public AcSpecsSheetComparer(string sheetName, string baseFile, string compFile, List<InstanceRow> instanceRows) : base(sheetName, baseFile, compFile, "Symbol", [])
        {
            UseList = new HashSet<string>([.. instanceRows.Select(x => x.AcCategory).Distinct()], StringExtensions.IgnoreCase);
            KeyColHeaders = ["Symbol"];
            SheetType = EnumSheetType.DTACSpecSheet;
        }

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

            Dictionary<string, List<Location>> dicBaseSheetItems = LoadSheetItems(BaseSheet, BaseSheetStartRow, BaseSheetEndRow, BaseSheetStartCol, BaseSheetEndCol);
            Dictionary<string, List<Location>> dicCompSheetItems = LoadSheetItems(CompSheet, CompSheetStartRow, CompSheetEndRow, CompSheetStartCol, CompSheetEndCol);

            GetCategoryIndex();
            BaseSheetVersion = SheetProvider.GetSheetVersion(BaseSheet);

            CompareItem(dicBaseSheetItems, dicCompSheetItems, out int totalCount, out int failCount, out int semiAutoCount, out string message);

            IEnumerable<string> lines = BaseSheet.Select(row => string.Join("\t", row));
            File.WriteAllLines(BaseFile, lines);

            result.SheetName = SheetName;
            result.SheetType = SheetType;
            result.TotalCount = totalCount;
            result.ManualCount = failCount;
            result.SemiAutoCount = semiAutoCount;
            result.AddDescription(message);
            return [result];
        }

        private void CompareItem(Dictionary<string, List<Location>> dicBaseSheetItems, Dictionary<string, List<Location>> dicCompSheetItems, out int totalCount, out int failCount, out int semiAutoCount, out string message)
        {
            totalCount = 0;
            failCount = 0;
            message = "";
            var failedCategories = new HashSet<string>();
            var missingCategories = new HashSet<string>();
            int startRow = BaseSheetVersion.EqualsIgnoreCase("2.0") ? 3 : 1;
            int categoryCountPerSymbol = _baseDicCategoryColIndex.Keys.Intersect(UseList).Count();
            foreach (string key in dicBaseSheetItems.Keys)
            {
                Location? location = SheetProvider.GetFirstLocation(dicBaseSheetItems, key);
                if (location != null)
                {
                    int row = location.Row;
                    totalCount += categoryCountPerSymbol;
                    if (dicCompSheetItems.TryGetValue(key, out List<Location>? locations))
                    {
                        Location? compLocation = locations.FirstOrDefault(l => !l.IsUsed);
                        if (compLocation == null)
                        {
                            failedCategories.AddRange(_baseDicCategoryColIndex.Keys);
                            SetCellSafe(BaseSheet, row, 0, "F");
                        }
                        else
                        {
                            foreach (string categoryName in _baseDicCategoryColIndex.Keys)
                            {
                                if (!UseList.Contains(categoryName))
                                {
                                    continue;
                                }
                                int baseCategoryStartCol = _baseDicCategoryColIndex[categoryName];
                                if (!_compareDicCategoryColIndex.TryGetValue(categoryName, out int compCategoryStartCol))
                                {
                                    failedCategories.Add(categoryName);
                                    missingCategories.Add(categoryName);
                                    SetCellSafe(BaseSheet, 0, baseCategoryStartCol, "Missing");
                                }
                                else
                                {
                                    for (int i = 0; i < startRow; i++)
                                    {
                                        if (row + i > BaseSheetEndRow)
                                        {
                                            continue;
                                        }

                                        for (int j = 0; j < 3; j++)
                                        {
                                            if (baseCategoryStartCol + j > BaseSheetEndCol)
                                            {
                                                continue;
                                            }

                                            if (compCategoryStartCol + j > CompSheetEndCol)
                                            {
                                                continue;
                                            }
                                            string baseContext = GetCellSafe(BaseSheet, row + i, baseCategoryStartCol + j);
                                            string comContext = GetCellSafe(CompSheet, compLocation.Row, compCategoryStartCol + j);

                                            bool basValue = double.TryParse(baseContext.TrimStart('=').Trim(), out double value1);
                                            if (basValue)
                                            {
                                                bool comValue = double.TryParse(comContext.TrimStart('=').Trim(), out double value2);
                                                if (comValue && Math.Abs(value1 - value2) < Epsilon)
                                                {
                                                    continue;
                                                }
                                            }

                                            if (!baseContext.Trim('=').EqualsIgnoreCase(comContext.Trim('=')))
                                            {
                                                SetCellSafe(BaseSheet, row, baseCategoryStartCol + j, $"{baseContext} => {comContext}");
                                                failedCategories.Add(categoryName);
                                            }
                                        }
                                    }
                                }
                            }
                            compLocation.IsUsed = true;
                        }
                    }
                    else
                    {
                        failedCategories.AddRange(_baseDicCategoryColIndex.Keys);
                        SetCellSafe(BaseSheet, row, 0, "F");
                    }
                }
            }

            totalCount = _baseDicCategoryColIndex.Keys.Intersect(UseList.Where(u => !string.IsNullOrEmpty(u)), StringExtensions.IgnoreCase).Count(k => !string.IsNullOrEmpty(k));
            failCount = failedCategories.Intersect(UseList.Where(u => !string.IsNullOrEmpty(u)), StringExtensions.IgnoreCase).Count(k => !string.IsNullOrEmpty(k));
            semiAutoCount = 0;
            var messages = new List<string>();
            if (missingCategories.Count != 0)
            {
                messages.Add($"Missing Category :\n{string.Join(",", missingCategories)}");
            }
            if (failedCategories.Count != 0)
            {
                messages.Add($"Failed Category :\n{string.Join(",", failedCategories)}");
            }
            if (messages.Count != 0)
            {
                string text = string.Join(";", messages);
                const int maxLength = 8000;
                if (string.IsNullOrEmpty(text) || text.Length < maxLength)
                {
                    message = text;
                }
                else
                {
                    message = string.Concat(text.AsSpan(0, maxLength - 3), "...");
                }
            }
        }

        private void GetCategoryIndex()
        {
            _baseDicCategoryColIndex = new Dictionary<string, int>(StringExtensions.IgnoreCase);
            _compareDicCategoryColIndex = new Dictionary<string, int>(StringExtensions.IgnoreCase);
            string categoryName;
            bool findCategoryRow = false;

            for (int row = 1; row <= BaseSheetEndRow; row++)
            {
                for (int col = 1; col <= BaseSheetEndCol; col++)
                {
                    string baseContext = GetCellSafe(BaseSheet, row, col);
                    if (Reg.RegexSelector().IsMatch(baseContext))
                    {
                        findCategoryRow = true;
                        col++;
                        while (col <= BaseSheetEndCol)
                        {
                            categoryName = GetCellSafe(BaseSheet, row, col).ToUpper();
                            if (!string.IsNullOrEmpty(categoryName) && !_baseDicCategoryColIndex.ContainsKey(categoryName))
                            {
                                _baseDicCategoryColIndex.Add(categoryName, col);
                            }

                            col++;
                        }
                        break;
                    }
                }
                if (findCategoryRow)
                {
                    break;
                }
            }

            findCategoryRow = false;
            for (int row = 1; row <= CompSheetEndRow; row++)
            {
                for (int col = 1; col <= CompSheetEndCol; col++)
                {
                    string value = GetCellSafe(CompSheet, row, col);
                    if (Reg.RegexSelector().IsMatch(value))
                    {
                        findCategoryRow = true;
                        col++;
                        while (col <= CompSheetEndCol)
                        {
                            categoryName = GetCellSafe(CompSheet, row, col);
                            if (!string.IsNullOrEmpty(categoryName) && !_compareDicCategoryColIndex.ContainsKey(categoryName))
                            {
                                _compareDicCategoryColIndex.Add(categoryName, col);
                            }

                            col++;
                        }
                        break;
                    }
                }
                if (findCategoryRow)
                {
                    break;
                }
            }
        }
    }
}
