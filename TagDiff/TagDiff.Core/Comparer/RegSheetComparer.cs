using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;

using TagDiff.Core.Common;
using TagDiff.Core.Output;

namespace TagDiff.Core.Comparer
{
    internal class RegSheetComparer : SheetComparerBase
    {
        public RegSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "", [])
        {
            SheetType = EnumSheetType.RegAssign;
        }

        public override List<CompareReport> Compare()
        {
            var groupedBase = new List<RegAssign>();
            var groupedCompare = new List<RegAssign>();

            groupedBase.AddRange(CombineSheetToList(BaseSheet, SheetName, true));
            groupedCompare.AddRange(CombineSheetToList(CompSheet, SheetName));
            (List<RegAssign> regAssigns, List<CompareReport> reports) = FindDifferences(groupedBase, groupedCompare);

            WriteRegAssignToTxt(regAssigns, BaseFile);

            var aggregatedResults = reports
                .GroupBy(x => x.SheetName)
                .Select(x => new CompareReport
                {
                    SheetType = EnumSheetType.RegAssign,
                    SheetName = x.Key,
                    ManualCount = x.Sum(ct => ct.ManualCount),
                    SemiAutoCount = x.Sum(ct => ct.SemiAutoCount),
                    TotalCount = x.Sum(ct => ct.TotalCount),
                    PassCount = x.Sum(ct => ct.PassCount)
                })
                .ToList();
            return aggregatedResults;
        }

        public static List<RegAssign> CombineSheetToList(List<string[]> sheet, string sheetName, bool fromBase = false)
        {
            var regAssign = new RegAssign();
            var setting = new Setting();
            bool isSettingCol = false;

            if (sheet.Count == 0)
            {
                return [];
            }

            int baseSheetEndRow = sheet.Count - 1;
            int baseSheetEndCol = sheet.Max(x => x.Length) - 1;
            var groupedData = new List<RegAssign>();
            for (int col = 0; col <= baseSheetEndCol; col++)
            {
                for (int row = 0; row <= baseSheetEndRow; row++)
                {
                    string cellString = GetCellSafe(sheet, row, col);
                    regAssign.SheetName = Path.GetFileNameWithoutExtension(sheetName);
                    if (row == 0)
                    {
                        string header = GetCellSafe(sheet, 0, col);
                        if (!string.IsNullOrEmpty(header))
                        {
                            isSettingCol = false;
                            groupedData.Add(regAssign);
                            regAssign = new RegAssign();
                            setting = new Setting();
                            if (cellString != null)
                            {
                                regAssign.Type = cellString;
                                regAssign.SheetName = Path.GetFileNameWithoutExtension(sheetName);
                                setting.ColNum = col;
                            }
                        }
                    }
                    else if (cellString != null)
                    {
                        if (row == 1)
                        {
                            setting.Header = cellString;
                            setting.ColNum = col;
                        }
                        else if (row == 2 && !isSettingCol)
                        {
                            string value = GetCellSafe(sheet, 2, col);
                            if (!string.IsNullOrEmpty(value))
                            {
                                regAssign.SettingName = cellString;
                                isSettingCol = true;
                            }
                        }
                    }

                    if (cellString != null && row >= 2)
                    {
                        if (cellString != regAssign.SettingName)
                        {
                            setting.Content.Add(cellString);
                        }
                    }
                    if (row == baseSheetEndRow)
                    {
                        setting.IsFromBase = fromBase;
                        regAssign.RegSettingList.Add(setting);
                        setting = new Setting();
                    }
                }
            }

            groupedData.Add(regAssign);
            return groupedData;
        }

        public static (List<RegAssign>, List<CompareReport>) FindDifferences(List<RegAssign> baseRegAssignList, List<RegAssign> comparedRegAssignList)
        {
            var groupedResult = new List<RegAssign>();
            var countRows = new List<CompareReport>();
            foreach (RegAssign baseRegAssign in baseRegAssignList)
            {
                var countRow = new CompareReport();
                RegAssign result = baseRegAssign;
                RegAssign assign = baseRegAssign;
                RegAssign? target = comparedRegAssignList.FirstOrDefault(x => x.SettingName.EqualsIgnoreCase(assign.SettingName) && x.Type.EqualsIgnoreCase(assign.Type));
                if (target != null)
                {
                    (result.RegSettingList, countRow) = FindSettingDifferences(baseRegAssign.RegSettingList, target.RegSettingList, result);
                    countRow.SheetName = baseRegAssign.SheetName;
                    countRow.SheetType = EnumSheetType.RegAssign;
                }
                else
                {
                    result.Status = "Missing";
                    foreach (Setting item in baseRegAssign.RegSettingList)
                    {
                        countRow.SheetName = baseRegAssign.SheetName;
                        countRow.SheetType = EnumSheetType.RegAssign;
                        if (item.Header != "SubBlockName" & item.Header != "InstanceName" & !string.IsNullOrEmpty(item.Header))
                        {
                            countRow.TotalCount += item.Content.Count;
                        }
                    }
                }
                if (baseRegAssign.RegSettingList.Count != 0)
                {
                    groupedResult.Add(result);
                    countRows.Add(countRow);
                }
            }

            return (groupedResult, countRows);
        }

        private static (List<Setting>, CompareReport countRow) FindSettingDifferences(List<Setting> baseSettingList, List<Setting> comparedSettingList, RegAssign regAssign)
        {
            var countRow = new CompareReport();
            regAssign.Status = "Pass";

            if (baseSettingList.Count < comparedSettingList.Count)
            {
                baseSettingList.AddRange(Enumerable.Repeat(new Setting(), comparedSettingList.Count - baseSettingList.Count));
                regAssign.RegSettingList.AddRange(Enumerable.Repeat(new Setting(), comparedSettingList.Count - regAssign.RegSettingList.Count));
            }

            if (baseSettingList.Count > comparedSettingList.Count)
            {
                comparedSettingList.AddRange(Enumerable.Repeat(new Setting(), baseSettingList.Count - comparedSettingList.Count));
            }

            if (baseSettingList.Count > 1)
            {
                int index = -1;
                int maxLength = baseSettingList.Max(bs => bs.Content.Count);
                for (int i = 0; i < maxLength; i++)
                {
                    if (baseSettingList.All(bs => string.IsNullOrEmpty(bs.Content.ElementAtOrDefault(i))))
                    {
                        index = i;
                        break;
                    }
                }
                if (index != -1)
                {
                    foreach (Setting baseSetting in baseSettingList)
                    {
                        baseSetting.Content = [.. baseSetting.Content.Take(index)];
                    }
                }
            }
            if (comparedSettingList.Count > 1)
            {
                int index = -1;
                int maxLength = comparedSettingList.Max(bs => bs.Content.Count);
                for (int i = 0; i < maxLength; i++)
                {
                    if (comparedSettingList.All(bs => string.IsNullOrEmpty(bs.Content.ElementAtOrDefault(i))))
                    {
                        index = i;
                        break;
                    }
                }
                if (index != -1)
                {
                    foreach (Setting comparedSetting in comparedSettingList)
                    {
                        comparedSetting.Content = [.. comparedSetting.Content.Take(index)];
                    }
                }
            }
            int settingLoopIdx = 0;
            // ReSharper disable once ForCanBeConvertedToForeach
            for (int i = 0; i < baseSettingList.Count; i++)
            {
                Setting baseSetting = baseSettingList[i];
                Setting comparedSetting = comparedSettingList[settingLoopIdx];

                if (baseSetting.Content.Count > comparedSetting.Content.Count)
                {
                    comparedSetting.Content.AddRange(Enumerable.Repeat("", baseSetting.Content.Count - comparedSetting.Content.Count));
                }
                else if (baseSetting.Content.Count < comparedSetting.Content.Count)
                {
                    baseSetting.Content.AddRange(Enumerable.Repeat("", comparedSetting.Content.Count - baseSetting.Content.Count));
                    regAssign.RegSettingList[settingLoopIdx] = baseSetting;
                }

                var resultContentList = new List<string>();
                foreach ((string baseStr, string comparedStr) in baseSetting.Content.Zip(comparedSetting.Content, (a, b) => (a, b)))
                {
                    if (baseStr.TrimStart('\'').EqualsIgnoreCase(comparedStr.TrimStart('\'')))
                    {
                        resultContentList.Add(baseStr);
                        if (baseSetting.Header != "SubBlockName" & baseSetting.Header != "InstanceName" & !string.IsNullOrEmpty(baseSetting.Header))
                        {
                            countRow.PassCount++;
                            countRow.TotalCount++;
                        }
                    }
                    else
                    {
                        resultContentList.Add(string.Format($"{baseStr} => {comparedStr}"));
                        regAssign.Status = "Diff";
                        if (baseSetting.Header != "SubBlockName" & baseSetting.Header != "InstanceName" & !string.IsNullOrEmpty(baseSetting.Header))
                        {
                            countRow.TotalCount++;
                            countRow.ManualCount++;
                        }
                    }
                }

                regAssign.RegSettingList[settingLoopIdx].Content = resultContentList;
                settingLoopIdx++;
            }

            return (regAssign.RegSettingList, countRow);
        }

        public static void WriteRegAssignToTxt(List<RegAssign> regAssigns, string txtFile)
        {
            // Write header row
            int maxCol = regAssigns.SelectMany(x => x.RegSettingList).Max(x => x.ColNum);
            int maxRow = regAssigns.SelectMany(x => x.RegSettingList).Max(x => x.Content.Count);

            var lines = new List<string>();
            string firstLine = "";
            foreach (RegAssign regAssign in regAssigns)
            {
                if (!string.IsNullOrEmpty(regAssign.Type))
                {
                    firstLine += regAssign.Type;
                    firstLine += new string('\t', regAssign.RegSettingList.Count);
                }
            }
            lines.Add(firstLine);

            var secondLine = new List<string>();
            foreach (RegAssign regAssign in regAssigns)
            {
                foreach (Setting setting in regAssign.RegSettingList)
                {
                    secondLine.Add(setting.Header);
                }
            }
            lines.Add(string.Join("\t", secondLine));

            var colToSetting = new Dictionary<int, Setting>();
            foreach (RegAssign ra in regAssigns)
            {
                foreach (Setting s in ra.RegSettingList)
                {
                    colToSetting[s.ColNum] = s;
                }
            }

            for (int i = 0; i < maxRow; i++)
            {
                var line = new List<string>(maxCol);
                for (int j = 0; j < maxCol; j++)
                {
                    if (colToSetting.TryGetValue(j, out Setting? found) && found.Content.Count > i)
                    {
                        string text = found.Content[i].Length > 8000 ? found.Content[i][..8000] : found.Content[i];
                        line.Add(text);
                    }
                    else
                    {
                        line.Add("");
                    }
                }

                if (i == 0)
                {
                    foreach (RegAssign regAssign in regAssigns)
                    {
                        if (!string.IsNullOrEmpty(regAssign.SettingName))
                        {
                            int j = regAssign.RegSettingList[0].ColNum;
                            line[j] = regAssign.SettingName.Length > 8000 ? regAssign.SettingName[..8000] : regAssign.SettingName;
                        }
                    }
                }

                if (line.Any(x => !string.IsNullOrEmpty(x)))
                {
                    lines.Add(string.Join("\t", line) + "\t");
                }
            }

            File.WriteAllLines(txtFile, lines);
        }
    }
}
