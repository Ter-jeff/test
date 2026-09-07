using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Harvest;

namespace Automation.GenerateIgxl.Basic.Business
{
    internal static class HarvestCoreMappingChecker
    {
        internal static readonly Regex _regexPl = new Regex(@"_PL\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        internal static readonly Regex _regex1 = new Regex("_(LPB|CON)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        internal static readonly Regex _regex2 = new Regex("_(CH)_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        internal static readonly Regex _regex3 = new Regex("_(SSC|SSU)_", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        internal static readonly Dictionary<string, Func<MappingCoreRow, string>> _contentMap =
            new Dictionary<string, Func<MappingCoreRow, string>>(StringComparer.OrdinalIgnoreCase)
            {
                ["inpattern"] = i => i.InitPattern,
                ["plpattern"] = i => i.Pattern,
                ["corename/pingroup"] = i => i.CoreName,
                ["harvestflag"] = i => i.HarvestFlag,
                ["powersupply (multiple rails bincut search)"] = i => i.PowerSupply,
                ["comment"] = i => i.Comment,
            };

        internal static void HarvestCheck()
        {
            PatSetSheet patsetAll = TestProgram.IgxlWorkBk.GetPatSetsSheet(IgxlWorkBook.PatSetsAll, FolderStructure.DirPatSetsAll);

            CreateMappingCoreTable(patsetAll);

            PreCheckHarvPinGrp();

            PreCheckSsn();
        }

        private static void PreCheckHarvPinGrp()
        {
            if (TestPlanStatic.MappingHarvestingSheet == null || TestPlanStatic.ScanInstanceSheets == null)
            {
                return;
            }

            foreach (BinCutInstanceSheet sheet in TestPlanStatic.ScanInstanceSheets)
            {
                foreach (BinCutInstanceRow row in sheet.Rows)
                {
                    if (!string.IsNullOrEmpty(row.PatternPinGroup))
                    {
                        foreach (string pat in row.PayloadList.Where(p => !string.IsNullOrEmpty(p)).ToList())
                        {
                            if (!Regex.IsMatch(pat, @"_PL\w{2}_", RegexOptions.IgnoreCase))
                            {
                                continue;
                            }
                            if (!TestPlanStatic.MappingHarvestingSheet.Rows.Exists(x => Regex.IsMatch(pat, "@" + x.Pattern, RegexOptions.IgnoreCase)))
                            {
                                ErrorReportManager.AddError(HarvestErrorType.W_MismatchPattern_01, sheet.SheetName, row.RowNum, 0, [pat]);
                            }
                        }
                    }
                }
            }
        }

        internal static void AddNewMappingCoreRow(
            MappingCoreRow newRow,
            MappingCoreTable mappingCoreTable,
            List<string> existRows
        )
        {
            string rowString = string.Join(
                    "|",
                    new List<string>
                    {
            newRow.InitPattern,
            newRow.Pattern,
            newRow.CoreName,
            newRow.HarvestFlag,
            newRow.PowerSupply,
            newRow.Comment,
                    }
                )
                .ToUpper();
            if (!existRows.Contains(rowString))
            {
                existRows.Add(rowString);
                mappingCoreTable.Rows.Add(newRow);
            }
        }

        private static void ProcessSingleCoreSsnRow(
            MappingHarvestingRow row,
            Dictionary<string, Dictionary<string, List<string>>> matchItemDic,
            List<string> allPatternNames,
            MappingCoreTable mappingCoreTable,
            List<string> existRows
        )
        {
            var regexCore = new Regex(
                $"{row.CoreName.ToLower().Replace("*", "(.+)?")}",
                RegexOptions.Compiled
            );
            var regexPattern = new Regex(
                $"{row.Pattern.Replace("*", "(.+)?")}",
                RegexOptions.Compiled
            );
            Dictionary<string, List<string>> groupByCluster;
            if (matchItemDic.TryGetValue(row.Pattern, out Dictionary<string, List<string>> value))
            {
                groupByCluster = value;
            }
            else
            {
                var matchItemNames = allPatternNames
                    .Where(x => regexPattern.IsMatch(x))
                    .ToList();
                groupByCluster = matchItemNames
                    .GroupBy(x => x.Split('_')[5])
                    .ToDictionary(x => x.Key.ToLower(), y => y.ToList());
                matchItemDic.Add(row.Pattern, groupByCluster);
            }

            foreach (KeyValuePair<string, List<string>> cluster in groupByCluster)
            {
                string clusterName = cluster.Key;
                List<string> patterns = cluster.Value;
                var ssnCoreNames = new List<string>();
                string plpattern = "";
                if (row.Pattern.ContainsIgnoreCase(clusterName))
                {
                    var newRow = new MappingCoreRow
                    {
                        InitPattern = row.InitPattern,
                        Pattern = row.Pattern,
                        CoreName = row.CoreName,
                        HarvestFlag = row.HarvestFlag,
                        PowerSupply = row.PowerSupply,
                        Comment = row.Comment
                    };
                    mappingCoreTable.Rows.Add(newRow);
                    continue;
                }
                foreach (string pattern in patterns)
                {
                    if (!LocalSpecs.HardIpInfos.ContainsKey(pattern))
                    {
                        continue;
                    }

                    HardIpInfo hipInfo = LocalSpecs.HardIpInfos.GetHardIpInfo(pattern);

                    IEnumerable<string> ssnInfos = hipInfo.SsnCoreName
                        .Replace(" ", "")
                        .Split(',')
                        .ToList()
                        .Where(x => regexCore.IsMatch(x.ToLower()))
                        .ToList()
                        .Distinct();
                    foreach (string ssnInfo in ssnInfos)
                    {
                        if (!ssnCoreNames.Exists(x => x.Equals(ssnInfo, StringComparison.OrdinalIgnoreCase)))
                        {
                            ssnCoreNames.Add(ssnInfo);
                        }
                    }
                    if (string.IsNullOrEmpty(plpattern))
                    {
                        List<string> patternKeywords = row.Pattern.Split('*').ToList();
                        patternKeywords.Add(clusterName);
                        var matchItems = new Dictionary<string, int>();
                        List<string> patternItems = pattern.Split('_').ToList();
                        int index = 0;
                        foreach (string item in patternItems)
                        {
                            index++;
                            if (!patternKeywords.Exists(x => x.ToLower().Equals(item.ToLower())))
                            {
                                continue;
                            }

                            matchItems.Add(item, index);
                        }
                        string prefix = matchItems.Values.Min() == 1 ? "" : "*";
                        string middle = string.Join("*", matchItems.OrderBy(x => x.Value)
                            .Select(x => x.Key)
                            .ToList());
                        string postfix = matchItems.Values.Max() == patternItems.Count ? "" : "*";
                        plpattern = $"{prefix}{middle}{postfix}";
                    }
                }
                if (!ssnCoreNames.Any())
                {

                    continue;
                }

                foreach (string ssnCoreName in ssnCoreNames)
                {
                    var newRow = new MappingCoreRow
                    {
                        InitPattern = row.InitPattern,
                        Pattern = plpattern,
                        CoreName = ssnCoreName,
                        HarvestFlag = row.HarvestFlag,
                        PowerSupply = row.PowerSupply,
                        Comment = row.Comment
                    };
                    AddNewMappingCoreRow(newRow, mappingCoreTable, existRows);
                }
            }
        }

        private static void CreateMappingCoreTable(PatSetSheet patsetAll)
        {
            var matchItemDic = new Dictionary<string, Dictionary<string, List<string>>>();
            if (patsetAll == null)
            {
                return;
            }
            if (TestPlanStatic.MappingHarvestingSheet == null)
            {
                return;
            }

            MappingHarvestingSheet mappingHarvTable = TestPlanStatic.MappingHarvestingSheet;
            var existRows = new List<string>();
            string sheetName = mappingHarvTable.SheetName;
            var mappingCoreTable = new MappingCoreTable(sheetName);
            var allPatternNames = patsetAll.Rows.Select(x => x.PatSetName).ToList();
            foreach (MappingHarvestingRow row in mappingHarvTable.Rows)
            {
                if (!row.IsSsn || row.CoreName.Split(';').Length > 1)
                {
                    var newRow = new MappingCoreRow
                    {
                        InitPattern = row.InitPattern,
                        Pattern = row.Pattern,
                        CoreName = row.CoreName,
                        HarvestFlag = row.HarvestFlag,
                        PowerSupply = row.PowerSupply,
                        Comment = row.Comment
                    };
                    AddNewMappingCoreRow(newRow, mappingCoreTable, existRows);
                }
                else
                {
                    ProcessSingleCoreSsnRow(
                        row,
                        matchItemDic,
                        allPatternNames,
                        mappingCoreTable,
                        existRows
                    );
                }
            }
            mappingCoreTable.CreateRegexList();
            ExcelWorksheet workSheet = EpWorkbook.TestPlanWorkbook.Worksheets.AddSheet(mappingCoreTable.SheetName);
            var titles = mappingHarvTable.HeaderIndex.Keys.ToList();
            int currentRow = 1;
            workSheet.Cells[currentRow++, 1].PrintExcelRow(titles.ToArray());
            foreach (MappingCoreRow row in mappingCoreTable.Rows)
            {
                var contents = new List<string>();
                foreach (KeyValuePair<string, int> content in mappingHarvTable.HeaderIndex)
                {
                    if (_contentMap.TryGetValue(content.Key, out Func<MappingCoreRow, string> selector))
                    {
                        contents.Add(selector(row));
                    }
                }
                workSheet.Cells[currentRow++, 1].PrintExcelRow(contents.ToArray());
            }
            Response.Report($"Adding {workSheet.Name} ...");
            workSheet.ExportWorkBook2Txt(FolderStructure.DirCommon);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, workSheet.Name);
            TestPlanStatic.MappingCoreTable = mappingCoreTable;
        }

        private static void PreCheckSsn()
        {
            PatSetSheet patSetAll = TestProgram.IgxlWorkBk.GetPatSetsSheet(IgxlWorkBook.PatSetsAll, FolderStructure.DirPatSetsAll);
            if (LocalSpecs.HardIpInfos == null || TestPlanStatic.ScanInstanceSheets == null || TestPlanStatic.MappingCoreTable == null || patSetAll == null)
            {
                return;
            }
            Dictionary<string, HardIpInfo> patInfoWithFullName = LocalSpecs.HardIpInfos.SelectMany(x => x.Value).ToDictionary(x => x.FullPattern, x => x, StringComparer.OrdinalIgnoreCase);

            MappingCoreTable mappingCoreTable = TestPlanStatic.MappingCoreTable;
            var scanPat = new List<string>();

            foreach (string pat in TestPlanStatic.ScanInstanceSheets.SelectMany(sheet => sheet.Rows.SelectMany(row => row.PayloadList.Where(pat => !scanPat.Exists(x => x.Equals(pat))))))
            {
                scanPat.Add(pat);
            }
            mappingCoreTable.AddToErrorReport();
            foreach (string pat in scanPat)
            {
                if (!_regexPl.IsMatch(pat) || _regex1.IsMatch(pat) || _regex2.IsMatch(pat))
                {
                    continue;
                }
                if (!_regex3.IsMatch(pat))
                {
                    continue;
                }
                PatSet foundPat = patSetAll.Rows.Find(x => x.PatSetName.Equals(pat, StringComparison.CurrentCultureIgnoreCase));
                if (foundPat == null)
                {
                    continue;
                }

                string fullPatName = foundPat.PatSetRows.FirstOrDefault()?.File.Split(':').Last();
                patInfoWithFullName.TryGetValue(fullPatName, out HardIpInfo hipInfo);
                if (hipInfo == null)
                {
                    continue;
                }
                List<string> ssnCores = hipInfo.SsnCoreName.Replace(" ", "").Split(',').ToList();
                List<MappingCoreRow> matchesItem = mappingCoreTable.Rows.FindAll(x => x.IsSsn() && Regex.IsMatch(pat, "@" + x.Pattern, RegexOptions.IgnoreCase));
                if (!matchesItem.Any())
                {
                    ErrorReportManager.AddError(SsnErrorType.W_MismatchPattern_02, "", 0, 0, $"Cannot found any matching pattern in the HarvestPinFlag_Table. Pattern : \"{hipInfo.FullPattern}\"", new string[] { hipInfo.FullPattern });
                    continue;
                }
                if (matchesItem.Count > 1)
                {
                    ErrorReportManager.AddError(SsnErrorType.W_MismatchPattern_01, "", 0, 0, $"Corresponding patterns > 1 in the HarvestPinFlag_Table. Pattern : \"{hipInfo.FullPattern}\"", new string[] { hipInfo.FullPattern });
                }
                MappingCoreRow matchItem = matchesItem.First();
                if (ssnCores.Count < 1)
                {
                    ErrorReportManager.AddError(SsnErrorType.E_MissingPatternInfo_01, "", 0, 0, $"Pattern : \"{hipInfo.FullPattern}\" , Missing SSN info in HardipInfo", new string[] { hipInfo.FullPattern });
                }
                var ssnCoreFromTable = matchItem.PinGrpFlagDic.Select(x => x.Key).ToList();
                foreach (string ssnCore in ssnCores)
                {
                    if (!ssnCoreFromTable.Exists(x => x.Equals(ssnCore, StringComparison.OrdinalIgnoreCase)))
                    {
                        ErrorReportManager.AddError(SsnErrorType.E_MismatchCore_01, "", 0, 0, $"Cannot found SsnCoreName \"{ssnCore}\" of pattern (HardipInfo) in the \"{matchItem.Pattern}\"(HarvestPinFlag_Table). Pattern : \"{hipInfo.FullPattern}\"", new string[] { ssnCore, $"{matchItem.Pattern}", $"{hipInfo.FullPattern}" });
                    }
                }
                foreach (string row in ssnCoreFromTable)
                {
                    if (!ssnCores.Exists(x => x.Equals(row, StringComparison.OrdinalIgnoreCase)))
                    {
                        ErrorReportManager.AddError(SsnErrorType.E_MismatchCore_02, "", 0, 0, $"Cannot found SsnCoreName \"{row}\" from \"{matchItem.Pattern}\" (HarvestPinFlag_Table) in the pattern (HardipInfo). Pattern : \"{hipInfo.FullPattern}\"", new string[] { row, matchItem.Pattern, hipInfo.FullPattern });
                    }
                }
            }
            foreach (MappingCoreRow row in mappingCoreTable.Rows.FindAll(x => x.IsSsn()))
            {
                foreach (KeyValuePair<string, string> dic in row.PinGrpFlagDic)
                {
                    if (dic.Value != "F_" + dic.Key)
                    {
                        ErrorReportManager.AddError(SsnErrorType.W_MismatchFlag_01, "", 0, 0, $"Pattern : \"{row.Pattern}\" , SsnCoreName : \"{dic.Key}\" SsnFlag : \"{dic.Value}\" mismatch", new string[] { row.Pattern, dic.Key, dic.Value });
                    }
                }

            }
        }
    }
}
