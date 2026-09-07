using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.Harvest;

namespace TestPlanLib.Utility
{
    internal static partial class HarvestingTruthTableReaderHelpers
    {
        private const string ConHarvestingTruthTable = "HarvestingTruthTable";
        private const string Regex1 = @"^Sum\((?<flag>.+)\)";
        private const string Regex2 = @"(?<cores>\[.+\])";
        private const string RegexOr = @"OR\((?<flag>.+)\)";

        [GeneratedRegex(RegexOr, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(Regex1, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(Regex2, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();

        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex4 = MyRegex2();

        private static void AddFlagDic(Flag flag, string flagName, Dictionary<string, List<string>> flagDic)
        {
            if (flag.Value != "x")
            {
                string flagTmp = flagName + "_" + flag.Value;
                if (flagDic.ContainsKey(flagName))
                {
                    if (!flagDic[flagName].Contains(flagTmp))
                    {
                        flagDic[flagName].Add(flagTmp);
                    }
                }
                else
                {
                    flagDic.Add(flagName, [flagTmp]);
                }
            }
        }

        internal static void AddFlags(ExcelWorksheet excelWorksheet, HeaderInfo headerInfo, string sheetName, HarvestingTruthTableSheet harvestingTruthTableSheet, int headerRowIndex, Dictionary<string, List<string>> sumFlagsDic, int i, HarvestingTruthTableRow harvestingTruthTableRow)
        {
            for (int j = headerInfo.Start; j <= headerInfo.End; j++)
            {
                string name = excelWorksheet.GetCellValue(headerRowIndex, j).Trim();
                string flagSyntax = "";
                string flagName = name;
                string enableStage = "";
                if (name.Contains("=>"))
                {
                    flagName = name.Split('>').First().Replace("=", "");
                    flagSyntax = name.Split('>').Last().Replace(" ", "");
                }
                if (flagName.Contains('@'))
                {
                    enableStage = flagName.Split('@').Last().Replace(" ", "");
                    flagName = flagName.Split('@').First().Replace(" ", "");
                }
                var sumFlags = new List<string>();
                var allFlags = new List<string>();
                bool isSum = false;
                bool isMultiGrp = false;
                bool isCust = false;

                HandleOrSum(harvestingTruthTableRow, flagSyntax, ref flagName, allFlags, ref isSum, ref isMultiGrp, ref isCust);

                if (!harvestingTruthTableRow.Flags.Exists(x => x.FlagName.EqualsIgnoreCase(flagName)) && !string.IsNullOrEmpty(flagName))
                {
                    var flag = new Flag
                    {
                        Name = name,
                        FlagName = flagName,
                        Value = excelWorksheet.GetCellValue(i, j).Replace("-", "To").Trim(),
                        IsSum = isSum,
                        IsCust = isCust
                    };
                    if (isCust || isSum)
                    {
                        AddFlagDic(flag, flagName, sumFlagsDic);
                    }
                    else
                    {
                        if (flag.Value.Contains("To"))
                        {
                            harvestingTruthTableSheet.AddError(HarvestErrorType.E_InvalidFormat_03, sheetName, i, j,
                                [excelWorksheet.GetCellValue(i, j), flag.Name]);
                        }
                    }
                    flag.IsMultiGroup = isMultiGrp;
                    flag.SumFlags = sumFlags;
                    flag.AllFlags = allFlags;
                    if (!string.IsNullOrEmpty(enableStage))
                    {
                        flag.EnableStage = enableStage;
                    }

                    harvestingTruthTableRow.Flags.Add(flag);
                }
                else if (!string.IsNullOrEmpty(flagName))
                {
                    int cnt = harvestingTruthTableRow.Flags.Count(x => x.FlagName.StartsWithIgnoreCase(flagName));
                    var flag = new Flag
                    {
                        Name = name,
                        FlagName = flagName + "_Part" + (cnt + 1),
                        Value = excelWorksheet.GetCellValue(i, j).Replace("-", "To").Trim(),
                        IsSum = isSum,
                        SumFlags = sumFlags,
                        AllFlags = allFlags
                    };
                    if (!string.IsNullOrEmpty(enableStage))
                    {
                        flag.EnableStage = enableStage;
                    }

                    harvestingTruthTableRow.Flags.Add(flag);
                }
            }
        }

        internal static string GetfusingValue(string address, bool isSingleAddress, string value, GroupItem groupItem, string groupValue)
        {
            string fusingValue = value;
            if (groupItem.Groups.Count > 0 && isSingleAddress)
            {
                string lastAddress = address.Split('>').Last();
                GroupItem iterates3 = lastAddress.GetIterates();
                if (iterates3.Groups.Count == 0)
                {
                    fusingValue = groupValue;
                }
            }

            return fusingValue;
        }

        internal static string GetJob(string sheetName, string commentTruthTableSheetName)
        {
            string job = "";
            if (Regex.IsMatch(sheetName, commentTruthTableSheetName, RegexOptions.IgnoreCase))
            {
                job = Regex.Replace(sheetName, commentTruthTableSheetName, "", RegexOptions.IgnoreCase).Trim().Trim('_');
            }
            else if (sheetName.StartsWithIgnoreCase(ConHarvestingTruthTable))
            {
                job = "CP1";
            }

            return job;
        }

        internal static string GetMergedCellValueWithGroup(ExcelWorksheet excelWorksheet, int rowNumber, int columnNumber, string value, Dictionary<int, List<Tuple<string, string>>> item)
        {
            string range = excelWorksheet.MergedCells[rowNumber, columnNumber];
            if (range != null)
            {
                int startColumn = new ExcelAddress(range).Start.Column;
                if (columnNumber - startColumn >= 0 && columnNumber - startColumn < item.Count)
                {
                    KeyValuePair<int, List<Tuple<string, string>>> group = item.ElementAt(columnNumber - startColumn);
                    foreach (Tuple<string, string> replace in group.Value)
                    {
                        value = value.Replace(replace.Item1, replace.Item2);
                    }

                    return value;
                }
            }
            return value;
        }

        internal static List<Fusing> GetReadFuses(ExcelWorksheet excelWorksheet, HeaderInfo headerInfo, int headerRowIndex, int i)
        {
            var readFuse = new List<Fusing>();
            for (int j = headerInfo.Start; j <= headerInfo.End; j++)
            {
                string field = excelWorksheet.GetCellValue(headerRowIndex - 1, j).Trim();
                string address = excelWorksheet.GetCellValue(headerRowIndex, j).Trim();
                bool isSingleAddress = string.IsNullOrEmpty(excelWorksheet.MergedCells[headerRowIndex, j]);
                var fusing = new Fusing { Field = field, Address = address };

                string value = excelWorksheet.GetCellValue(i, j).Trim();
                GroupItem iterates = value.GetIterates();
                string groupValue = GetMergedCellValueWithGroup(excelWorksheet, i, j, value, iterates.Groups).Trim();
                fusing.Value = GetfusingValue(address, isSingleAddress, value, iterates, groupValue);

                if (!string.IsNullOrEmpty(address))
                {
                    readFuse.Add(fusing);
                }
            }

            return readFuse;
        }

        internal static List<Flag> GetRepairFlags(ExcelWorksheet excelWorksheet, HeaderInfo headerInfo, int headerRowIndex, int i)
        {
            var repairFlags = new List<Flag>();
            for (int j = headerInfo.Start; j <= headerInfo.End; j++)
            {
                string name = excelWorksheet.GetCellValue(headerRowIndex, j).Trim();
                string flagName = name;
                string value = excelWorksheet.GetCellValue(i, j).Trim();
                var flag = new Flag { Name = name, FlagName = flagName, Value = value };
                if (!string.IsNullOrEmpty(name))
                {
                    repairFlags.Add(flag);
                }
            }

            return repairFlags;
        }

        private static void HandleOrSum(HarvestingTruthTableRow harvestingTruthTableRow, string flagSyntax, ref string flagName, List<string> allFlags, ref bool isSum, ref bool isMultiGrp, ref bool isCust)
        {
            #region OR
            if (_regex.IsMatch(flagName))
            {
                string flag = _regex.Match(flagName).Groups["flag"].ToString().Replace(")", "");
                string orFlagName = flag.Replace("+", "_").Replace(",", "_").Replace("[", "_").Replace("]", "_").Replace(":", "_").Replace("..", "_").Replace("__", "_");
                string[] sumflagItem = flag.Split([',', ' ', '+'], StringSplitOptions.RemoveEmptyEntries);
                if (sumflagItem.Length > 1)
                {
                    orFlagName = "";
                    foreach (string item in sumflagItem)
                    {
                        if (orFlagName?.Length != 0)
                        {
                            orFlagName += "OR_";
                        }

                        string flagTmp = _regex.IsMatch(item) ? _regex.Match(item).Groups["flag"].ToString() : item;
                        orFlagName += (flagTmp.Replace("[", "_").Replace("]", "_").Replace(":", "_").Replace("..", "_") + "_").Replace("__", "_");
                        GroupItem iterates = flagTmp.GetIterates();
                        if (iterates.Groups.Count != 0)
                        {
                            foreach (KeyValuePair<int, List<Tuple<string, string>>> iterate in iterates.Groups.OrderBy(x => x.Key))
                            {
                                string text = flagTmp;
                                foreach (Tuple<string, string> iterateItem in iterate.Value)
                                {
                                    text = text.Replace(iterateItem.Item1, iterateItem.Item2);
                                }

                                allFlags.Add(text);
                            }
                        }
                        else
                        {
                            allFlags.Add(item);
                        }
                    }
                    orFlagName += "SUM";
                    flagName = orFlagName;
                }
                else
                {
                    GroupItem iterates = flag.GetIterates();
                    foreach (KeyValuePair<int, List<Tuple<string, string>>> iterate in iterates.Groups)
                    {
                        string text = flag;
                        foreach (Tuple<string, string> item in iterate.Value)
                        {
                            text = text.Replace(item.Item1, item.Item2);
                        }

                        allFlags.Add(text);
                    }
                    flagName = orFlagName;
                }
                isMultiGrp = true;
                isCust = true;
            }
            #endregion

            #region SUM
            else if (_regex2.IsMatch(flagName))
            {
                string flag = _regex2.Match(flagName).Groups["flag"].ToString();
                string flagRemoveCode = _regex4.Replace(flag, "");
                string[] sumflagItem = flag.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries);
                string sumFlagName = (flag.Replace(",", "_").Replace("[", "_").Replace("]", "_").Replace(":", "_").Replace("..", "_") + "_SUM").Replace("__", "_");
                if (sumflagItem.Length > 1)
                {
                    sumFlagName = "";
                    foreach (string item in sumflagItem)
                    {
                        string flagTmp = _regex2.IsMatch(item) ? _regex2.Match(item).Groups["flag"].ToString() : item;
                        sumFlagName += (flagTmp.Replace("[", "_").Replace("]", "_").Replace(":", "_").Replace("..", "_") + "_").Replace("__", "_");
                        GroupItem iterates = flagTmp.GetIterates();
                        if (iterates.Groups.Count != 0)
                        {
                            foreach (KeyValuePair<int, List<Tuple<string, string>>> iterate in iterates.Groups.OrderBy(x => x.Key))
                            {
                                string text = flagTmp;
                                foreach (Tuple<string, string> iterateItem in iterate.Value)
                                {
                                    text = text.Replace(iterateItem.Item1, iterateItem.Item2);
                                }

                                allFlags.Add(text);
                            }
                        }
                        else
                        {
                            allFlags.Add(item);
                        }
                    }
                    sumFlagName += "SUM";
                    isMultiGrp = true;
                    flagName = sumFlagName;
                    isSum = true;
                }
                else
                {
                    GroupItem iterates = flag.GetIterates();
                    foreach (KeyValuePair<int, List<Tuple<string, string>>> iterate in iterates.Groups.OrderBy(x => x.Key))
                    {
                        string text = flag;
                        foreach (Tuple<string, string> item in iterate.Value)
                        {
                            text = text.Replace(item.Item1, item.Item2);
                        }

                        allFlags.Add(text);
                    }
                    flagName = sumFlagName;
                    isSum = true;
                }
                if (flagSyntax.Length != 0)
                {
                    //ex:(b0 , F_GFX_HARV_9_0_SUM_0) , use sum_0 flag to define group has harvest
                    harvestingTruthTableRow.FlagSyntaxDic.Add(flagSyntax, flagName + "_0");
                }
            }
            #endregion
            else
            {
                allFlags.Add(flagName);
            }
        }
    }
}
