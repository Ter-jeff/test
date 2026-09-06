using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using TagDiff.Core.Common;
using TagDiff.Core.Output;

namespace TagDiff.Core.Comparer
{
    public class TestInstanceSheetComparer : IgxlSheetComparerBase
    {
        private const string Cs = "_CS_";
        private readonly bool _onlyCompareCsharpInstance;

        public TestInstanceSheetComparer(string sheetName, string baseFile, string compFile, List<SubFlowSheet> subFlowSheets, bool onlyCompareCsharpInstance) : base(sheetName, baseFile, compFile, "Test Name", ["Arg0"])
        {
            _onlyCompareCsharpInstance = onlyCompareCsharpInstance;
            UseList = new HashSet<string>([.. subFlowSheets.SelectMany(x => x.Rows).Where(x => x.Opcode.EqualsIgnoreCase("Test")).Select(x => x.Parameter.Split(' ').First())], StringExtensions.IgnoreCase);
            KeyColHeaders = ["Test Name"];
            SheetType = EnumSheetType.DTTestInstancesSheet;
        }

        public override List<CompareReport> Compare()
        {
            var result = new CompareReport();
            if (CompSheet.Count == 0)
            {
                result.SheetName = SheetName;
                result.SheetType = EnumSheetType.DTTestInstancesSheet;
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
            var description = new List<string>();
            for (int row = BaseSheetStartRow + 1; row <= BaseSheetEndRow; row++)
            {
                string itemKey = GetItemKey(BaseSheet, row, baseKeyIndices).Trim();
                if (string.IsNullOrEmpty(itemKey) || !UseList.Contains(itemKey))
                {
                    continue;
                }

                if (_onlyCompareCsharpInstance)
                {
                    string type = GetCellSafe(BaseSheet, row, 2).TrimStart('\'').ToUpper();
                    if (type != ".NET")
                    {
                        continue;
                    }
                    if (itemKey.ContainsIgnoreCase(Cs))
                    {
                        itemKey = itemKey.Replace(Cs, "_");
                    }
                    if (itemKey.EndsWithIgnoreCase(Cs))
                    {
                        itemKey = itemKey.ReplaceEndsWith(Cs, "_");
                    }
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
                    else if (!CompareSheetRow(itemKey, new Location(row, BaseSheetStartCol), compLocation, Skips, ref description, _onlyCompareCsharpInstance))
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
                    failCount++;
                    SetCellSafe(BaseSheet, row, 0, "F");
                }
            }

            IEnumerable<string> lines = BaseSheet.Select(row => string.Join("\t", row));
            File.WriteAllLines(BaseFile, lines);

            result.SheetName = SheetName;
            result.SheetType = EnumSheetType.DTTestInstancesSheet;
            result.TotalCount = totalCount;
            result.ManualCount = failCount;
            result.SemiAutoCount = semiAutoCount;

            var descriptionDic = description.GroupBy(x => x).ToDictionary(x => x.Key, x => x.Count());
            if (descriptionDic.Count != 0)
            {
                string message = "SemiAuto: ";
                for (int i = 0; i < descriptionDic.Count; i++)
                {
                    KeyValuePair<string, int> pair = descriptionDic.ElementAt(i);
                    string str = pair.Key + "[" + pair.Value + "]";
                    if (i == 0)
                    {
                        message += str;
                    }
                    else
                    {
                        message += "," + str;
                    }
                }

                result.AddDescription(message);
            }

            return [result];
        }

        protected bool CompareSheetRow(string itemKey, Location baseLocation, Location compLocation, HashSet<int> skips, ref List<string> reasonList, bool ignoreCsTestName = false)
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

                string baseContext = (col < baseRow.Length ? baseRow[col] ?? string.Empty : string.Empty).TrimStart('\'');
                string comContext = (compRow != null && col < compRowLen ? compRow[col] ?? string.Empty : string.Empty).TrimStart('\'');
                string baseStripped = baseContext.Trim('=');
                string comStripped = comContext.Trim('=');
                if (!baseStripped.EqualsIgnoreCase(comStripped))
                {
                    if (!(double.TryParse(baseStripped, out double baseValue) && double.TryParse(comStripped, out double comValue) && Math.Abs(baseValue - comValue) < Epsilon))
                    {
                        //Skip compare TestName and Overlay when the string has "_CS_" for CS correlation TP.
                        if (col == baseLocation.Col || col == 14)
                        {
                            string temp = baseContext;
                            if (itemKey.ContainsIgnoreCase(Cs))
                            {
                                temp = itemKey.Replace(Cs, "_");
                            }
                            if (itemKey.EndsWithIgnoreCase(Cs))
                            {
                                temp = itemKey.ReplaceEndsWith(Cs, "_");
                            }
                            if (temp.EqualsIgnoreCase(comContext) && ignoreCsTestName)
                            {
                                SetCellSafe(BaseSheet, baseLocation.Row, col, $"{baseContext} => {comContext}");
                            }
                        }
                        else
                        {
                            UpdateFailReasonList(ref reasonList, col + 1);
                            SetCellSafe(BaseSheet, baseLocation.Row, col, $"{baseContext} => {comContext}");
                            result = false;
                        }
                    }
                }
            }
            return result;
        }

        private static void UpdateFailReasonList(ref List<string> reasonList, int col)
        {
            if (col == 4)
            {
                reasonList.Add("VBT Name");
            }
            else if (col == 6)
            {
                reasonList.Add("DC category");
            }
            else if (col == 7)
            {
                reasonList.Add("DC selector");
            }
            else if (col == 8)
            {
                reasonList.Add("AC category");
            }
            else if (col == 9)
            {
                reasonList.Add("AC selector");
            }
            else if (col == 10)
            {
                reasonList.Add("Time Sets");
            }
            else if (col == 11)
            {
                reasonList.Add("Edge Sets");
            }
            else if (col == 12)
            {
                reasonList.Add("Pin Levels");
            }
            else if (col == 13)
            {
                reasonList.Add("Mixed Signal Timing");
            }
            else if (col == 14)
            {
                reasonList.Add("Overlay");
            }
            else if (col > 15)
            {
                reasonList.Add("Arguments value");
            }
        }
    }
}
