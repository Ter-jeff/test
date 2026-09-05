using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.BistBira.Base;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.GenerateIgxl.PostAction.VreValidationReport;
using Automation.InputManager;
using Automation.Static;

using OfficeOpenXml;

using ScghLib.Base;
using ScghLib.Enums;
using ScghLib.Reader;
using ScghLib.Utility;

using TestPlanLib;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinNumber;
using TestPlanLib.Harvest;
using TestPlanLib.Singleton;
using TestPlanLib.Utility;

namespace Automation.GenerateIgxl.PostAction.GenVreTestCase
{
    public class GenVreTestCaseMain
    {
        private VreTestCaseTable _vreTestCaseTable;
        private MappingCoreTable _mappingCoreTable;

        private VreMbistLookupTable _vreMbistLookupTable;
        private HarvestingTruthTableSheet _harvestingTruthTable;
        private List<VreTestCaseRow> _vreTestCaseTableLevel2 = new List<VreTestCaseRow>();
        private List<FlagOperationSheet> _flagOperationSheets;
        private Dictionary<string, BistInfo> _bistInfoDic;
        private HashSet<string> _bistHarvPatterns = new(StringComparer.OrdinalIgnoreCase);
        private HashSet<string> _scanHarvPatterns = new(StringComparer.OrdinalIgnoreCase);
        public GenVreTestCaseMain(VreTestCaseTable vreTestCaseTable, MappingCoreTable mappingCoreTable, HarvestingTruthTableSheet harvestingTruthTable, VreMbistLookupTable vreMbistLookupTable, List<FlagOperationSheet> flagOperationSheets)
        {
            _vreTestCaseTable = vreTestCaseTable;
            _mappingCoreTable = mappingCoreTable;
            _harvestingTruthTable = harvestingTruthTable;
            _vreMbistLookupTable = vreMbistLookupTable;
            _flagOperationSheets = flagOperationSheets;
        }
        public void Workflow()
        {
            LocalSpecs.ElapsedTimeResults.MeasureSection("VreTestCase", () =>
            {
                GenTestCaseFromDigCores();
                GenTestCaseFromMbistLookup();
                MergeLevel2Case();
                Write();
            });
            LocalSpecs.ElapsedTimeResults.MeasureSection("VreValidationReport", () =>
            {
                new VreValidationReportMain(_bistHarvPatterns, _scanHarvPatterns).WorkFlow();
            });
        }

        public void Write()
        {
            List<string> titles = _vreTestCaseTable.HeaderIndex.Keys.ToList();
            string outputPath = Path.Combine(FolderStructure.DirIgLink, $"{_vreTestCaseTable.SheetName}.csv");
            using (var writer = new StreamWriter(outputPath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine(string.Join(",", titles));
                foreach (VreTestCaseRow row in _vreTestCaseTable.Rows)
                {
                    var contents = new List<string>();
                    foreach (KeyValuePair<string, int> content in _vreTestCaseTable.HeaderIndex)
                    {
                        if (content.Key.ToLower() == "case id")
                        {
                            contents.Add(row.CaseId.ToString());
                        }
                        else if (content.Key.ToLower() == "subprogram")
                        {
                            contents.Add(row.SubProgram);
                        }
                        else if (content.Key.ToLower() == "procesurename")
                        {
                            contents.Add(row.ProcesureName);
                        }
                        else if (content.Key.ToLower() == "instancename")
                        {
                            contents.Add(row.InstanceName);
                        }
                        else if (content.Key.ToLower() == "pattern1" && row.Pattern.Count() > 0)
                        {
                            contents.Add(row.Pattern[0]);
                        }
                        else if (content.Key.ToLower() == "pattern2" && row.Pattern.Count() > 1)
                        {
                            contents.Add(row.Pattern[1]);
                        }
                        else if (content.Key.ToLower() == "pattern3" && row.Pattern.Count() > 2)
                        {
                            contents.Add(row.Pattern[2]);
                        }
                        else if (content.Key.ToLower() == "user_def")
                        {
                            contents.Add(row.UserDef);
                        }
                        else if (content.Key.ToLower() == "expected hard bin")
                        {
                            contents.Add(row.HardBin);
                        }
                        else if (content.Key.ToLower() == "expected soft bin")
                        {
                            contents.Add(row.SorfdBin);
                        }
                        else if (content.Key.ToLower() == "comment")
                        {
                            contents.Add(row.Comment);
                        }
                        else if (content.Key.ToLower() == "level check")
                        {
                            contents.Add(row.LevelCheck);
                        }
                        else
                        {
                            contents.Add("");
                        }
                    }
                    writer.WriteLine(string.Join(",", contents));
                }
            }
        }

        private void MergeLevel2Case()
        {
            int currentCaseId = _vreTestCaseTable.GetMaxCaseId + 1;
            foreach (VreTestCaseRow row in _vreTestCaseTableLevel2)
            {
                row.CaseId = currentCaseId++;
                _vreTestCaseTable.Rows.Add(row);
            }
        }


        private void GenTestCaseFromMbistLookup()
        {
            string functionName = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.CSharpFuncNameFuncTestMain, "scan", true).FullFunctionName;
            if (_vreMbistLookupTable == null)
            {
                return;
            }

            var mbistSheets = new List<MbistSheet>();
            int currentCaseId = _vreTestCaseTable.GetMaxCaseId + 1;
            if (TestPlanStatic.MainFlowSheet != null && TestPlanStatic.MainFlowSheet.Rows != null)
            {
                List<FlowSequenceNew> mbistFlows = TestPlanStatic.MainFlowSheet.Rows.FirstOrDefault().SequencesNew.FindAll(x => x.Module.ToUpper().Equals("MBIST"));
                mbistSheets = new BistBiraInputManager(EpWorkbook.TestPlanWorkbook, null).AddSourceSheetFromFlowMain(mbistFlows);
            }
            var prodFlowSheets = new List<BistProdFlowSheet>();
            foreach (MbistSheet mbistSheet in mbistSheets)
            {
                string sheetName = mbistSheet.SheetName;
                ExcelWorksheet worksheet = EpWorkbook.ScghWorkbook.Worksheets[sheetName];
                if (worksheet != null)
                {
                    var sheetReader = new BistProdFlowReader(mbistSheet);
                    prodFlowSheets.Add(sheetReader.ReadSheet(worksheet));
                }
            }
            _bistInfoDic = GetBistInformation(prodFlowSheets);
            foreach (OreMbistLookupRow row in _vreMbistLookupTable.Rows)
            {
                foreach (string serverName in row.Servers)
                {
                    foreach (string groupName in row.MemoryGroups)
                    {
                        foreach (string excludePattern in row.ExcludePatterns)
                        {
                            List<KeyValuePair<string, BistInfo>> targetPatterns = _bistInfoDic.ToList().FindAll(x => x.Key.Contains(serverName) && x.Key.Contains(groupName) && (string.IsNullOrEmpty(excludePattern) || !x.Key.Contains(excludePattern)) && x.Key.Split('_')[9].StartsWith(row.Pmode));
                            foreach (KeyValuePair<string, BistInfo> targetPattern in targetPatterns)
                            {
                                string hardBin = GetExpectedHardBin(targetPattern.Value.HarvsetFlag);
                                if (!_bistHarvPatterns.Contains(targetPattern.Key))
                                {
                                    _bistHarvPatterns.Add(targetPattern.Key);
                                }
                                VreTestCaseRow testCaseRow = new VreTestCaseRow { CaseId = (currentCaseId++), SubProgram = "Sub_Digital", Pattern = new List<string> { targetPattern.Key }, InstanceName = "*", LevelCheck = "1", HardBin = hardBin, SorfdBin = "*", ProcesureName = functionName, Comment = hardBin == "Not found" ? targetPattern.Value.HarvsetFlag : "" };
                                _vreTestCaseTable.Rows.Add(testCaseRow);
                                foreach (string level in targetPattern.Value.Levels)
                                {
                                    VreTestCaseRow testCaseRowLevel = new VreTestCaseRow { SubProgram = "Sub_Digital", Pattern = new List<string> { targetPattern.Key }, InstanceName = $"*_{level}", LevelCheck = "2", HardBin = hardBin, SorfdBin = "*", ProcesureName = functionName };
                                    _vreTestCaseTableLevel2.Add(testCaseRowLevel);
                                }
                            }
                        }
                    }
                }
            }
        }
        private Dictionary<string, BistInfo> GetBistInformation(List<BistProdFlowSheet> prodFlowSheets)
        {
            Dictionary<string, BistInfo> bistInfoDic = new Dictionary<string, BistInfo>();
            foreach (BistProdFlowSheet sheet in prodFlowSheets)
            {
                var labelHarvFlagDic = sheet.Rows.Where(x => x.BistActionType == BistActionType.Set).ToDictionary(x => x.Label, BistAction.SetActionParameter);
                foreach (BistProdFlowRow row in sheet.Rows)
                {
                    if (row.BistActionType != BistActionType.RunPattern)
                    {
                        continue;
                    }
                    string pattern = row.Pattern;
                    string targetLabel = row.FailBranch;
                    string harvestFlag = labelHarvFlagDic.ContainsKey(targetLabel) ? labelHarvFlagDic[targetLabel] : "";
                    if (string.IsNullOrEmpty(harvestFlag))
                    {
                        continue;
                    }
                    string level = row.IsDsscRow ? BistNaming.GetVoltageType(row.OriVoltage) : BistNaming.GetVoltageType(row.Voltage);
                    BistInfo bistInfo = new BistInfo { HarvsetFlag = harvestFlag, Levels = new List<string> { level } };
                    if (bistInfoDic.ContainsKey(pattern))
                    {
                        bistInfo = bistInfoDic[pattern];
                        bistInfo.HarvsetFlag = harvestFlag;
                        if (!bistInfo.Levels.Contains(level))
                        {
                            bistInfo.Levels.Add(level);
                        }
                    }
                    else
                    {
                        bistInfoDic.Add(pattern, bistInfo);
                    }
                }
            }
            return bistInfoDic;
        }
        private void GenTestCaseFromDigCores()
        {
            int currentCaseId = _vreTestCaseTable.GetMaxCaseId + 1;
            Dictionary<string, HashSet<string>> patLevelDic = GetPatternWIthLevel();
            string functionName = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.CSharpFuncNameFuncTestMain, "scan", true).FullFunctionName;
            string binCutFunctionName = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.CSharpFuncNameBinCutTest, "bincut", true).FullFunctionName;
            Dictionary<string, List<string>> usedPatternLevelDic = new Dictionary<string, List<string>>();
            foreach (MappingCoreRow row in _mappingCoreTable.Rows)
            {
                List<string> harvetFlags = row.HarvestFlag.Split(';').ToList();
                int flagIndex = 0;
                foreach (string coreName in row.CoreName.Split(';'))
                {
                    string userDef = (row.IsSsn() ? "FailCore:" : "FailPin:") + coreName;
                    string harvestFlag = harvetFlags[flagIndex++];
                    string hardBin = GetExpectedHardBin(harvestFlag);
                    string pattern = row.Pattern;
                    HashSet<string> procesureNames = new(StringComparer.OrdinalIgnoreCase);
                    Regex patRegex = new Regex($"{row.Pattern.Replace("*", "(.+)?")}", RegexOptions.Compiled);
                    if (!usedPatternLevelDic.TryGetValue(pattern, out List<string>? allLevels))
                    {
                        IEnumerable<KeyValuePair<string, HashSet<string>>> matchPatterns = patLevelDic.Where(x => patRegex.IsMatch(x.Key));
                        foreach (KeyValuePair<string, HashSet<string>> matchPattern in matchPatterns)
                        {
                            if (_scanHarvPatterns.Contains(matchPattern.Key))
                            {
                                continue;
                            }
                            _scanHarvPatterns.Add(matchPattern.Key);
                        }
                        if (matchPatterns != null && matchPatterns.Any())
                        {
                            allLevels = matchPatterns.SelectMany(x => x.Value).Distinct().ToList();
                            if (allLevels.Any())
                            {
                                usedPatternLevelDic.Add(pattern, allLevels);
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                    foreach (string level in allLevels)
                    {
                        string procesureName = level.Contains("BV") ? binCutFunctionName : functionName;
                        _vreTestCaseTableLevel2.Add(new VreTestCaseRow
                        {
                            SubProgram = "Sub_Digital",
                            InstanceName = $"*_{level}",
                            ProcesureName = procesureName,
                            UserDef = userDef,
                            Pattern = new List<string> { pattern },
                            HardBin = hardBin,
                            SorfdBin = "*",
                            LevelCheck = "2",
                            Comment = row.Comment
                        });
                        if (!procesureNames.Contains(procesureName))
                        {
                            procesureNames.Add(procesureName);
                        }
                    }
                    foreach (string procesureName in procesureNames)
                    {
                        _vreTestCaseTable.Rows.Add(new VreTestCaseRow
                        {
                            CaseId = (currentCaseId++),
                            SubProgram = "Sub_Digital",
                            InstanceName = "*",
                            ProcesureName = procesureName,
                            UserDef = userDef,
                            Pattern = new List<string> { pattern },
                            HardBin = hardBin,
                            SorfdBin = "*",
                            LevelCheck = "1",
                            Comment = hardBin == "Not found" ? harvestFlag : "",
                        });
                    }
                }
            }
        }
        private Dictionary<string, HashSet<string>> GetPatternWIthLevel()
        {
            Dictionary<string, HashSet<string>> patLevelDic = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<BinCutInstanceSheet>> allInstanceSheet = new Dictionary<string, List<BinCutInstanceSheet>>();
            allInstanceSheet.Add("Scan", TestPlanStatic.ScanInstanceSheets);
            allInstanceSheet.Add("Bincut", TestPlanStatic.BinCutInstanceSheets);

            foreach (KeyValuePair<string, List<BinCutInstanceSheet>> type in allInstanceSheet)
            {
                bool isBincut = type.Key == "Bincut";
                List<BinCutInstanceSheet> sheets = type.Value;
                foreach (BinCutInstanceSheet sheet in sheets)
                {
                    foreach (BinCutInstanceRow row in sheet.Rows)
                    {
                        string level = GetLevel(row, isBincut);
                        if (string.IsNullOrEmpty(level))
                        {
                            continue;
                        }

                        foreach (string pattern in row.PatternList)
                        {
                            if (patLevelDic.ContainsKey(pattern))
                            {
                                HashSet<string> targetLevelList = patLevelDic[pattern];
                                if (!targetLevelList.Contains(level))
                                {
                                    targetLevelList.Add(level);
                                }
                            }
                            else
                            {
                                patLevelDic.Add(pattern, new HashSet<string> { level });
                            }
                        }
                    }
                }
            }
            return patLevelDic;
        }
        private string GetLevel(BinCutInstanceRow row, bool isBincut)
        {
            if (isBincut)
            {
                return row.SubFlow.ToUpper().Contains("HBV") ? "HBV" : "BV";
            }
            else
            {
                string level = GetVoltageType(row);
                if (!level.Equals("UnknowType"))
                {
                    return level;
                }
                else
                {
                    return "";
                }
            }
        }
        public string GetVoltageType(BinCutInstanceRow row)
        {
            string typeByDc = BinCutInstanceRowUtility.GetTypeByFlowNameOrDcCategory(row.DCcategory);
            if (!typeByDc.Equals("UnknowType"))
            {
                return typeByDc;
            }

            string typeByFlow = BinCutInstanceRowUtility.GetTypeByFlowNameOrDcCategory(row.FlowName);
            if (!typeByFlow.Equals("UnknowType"))
            {
                return typeByFlow;
            }

            return "UnknowType";
        }
        private string GetExpectedHardBin(string harvFlag)
        {
            if (_harvestingTruthTable == null)
            {
                return "";
            }
            HashSet<string> enableFlags = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { harvFlag };

            ExcuteFlagOperation(enableFlags);

            foreach (HarvestingTruthTableRow row in _harvestingTruthTable.Rows)
            {
                bool match = true;

                foreach (Flag flag in row.Flags)
                {
                    if (flag.Value.ToUpper() == "X")
                    {
                        continue;
                    }

                    if (flag.IsSum)
                    {
                        int.TryParse(flag.Value.Replace("To", "-").Split('-').FirstOrDefault(), out int min);
                        int.TryParse(flag.Value.Replace("To", "-").Split('-').LastOrDefault(), out int max);
                        int enableCounts = flag.AllFlags.Count(enableFlags.Contains);
                        if (enableCounts < min || enableCounts > max)
                        {
                            match = false;
                            break;
                        }
                    }
                    else
                    {
                        if (flag.Value == "1" && !enableFlags.Contains(flag.FlagName))
                        {
                            match = false;
                            break;
                        }
                        if (flag.Value == "0" && enableFlags.Contains(flag.FlagName))
                        {
                            match = false;
                            break;
                        }
                    }
                }
                if (match)
                {
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Harvest", row.Condition, "", null, true);
                    return binNumInfo.BinNumInfo.HardBin.ToString();
                }
            }
            return "Not found";
        }
        private void ExcuteFlagOperation(HashSet<string> enableFlags)
        {
            foreach (FlagOperationSheet sheet in _flagOperationSheets)
            {
                ExceuteSummary(enableFlags);
                foreach (FlagOperationRow row in sheet.Rows)
                {
                    int enableCount = row.ExistFlags.Count(x => (x.Contains("!") && !enableFlags.Contains(x.Replace("!", ""))) || (!x.Contains("!") && enableFlags.Contains(x)));
                    if (enableFlags.Contains(row.NewFlag))
                    {
                        continue;
                    }
                    if (row.Operator == "OR" && enableCount > 0)
                    {
                        enableFlags.Add(row.NewFlag);
                        continue;
                    }
                    if (row.Operator == "EQUAL" && enableCount == 1)
                    {
                        enableFlags.Add(row.NewFlag);
                        continue;
                    }
                    if (row.Operator == "AND" && enableCount == row.ExistFlags.Count())
                    {
                        enableFlags.Add(row.NewFlag);
                    }
                }
            }
        }
        private void ExceuteSummary(HashSet<string> enableFlags)
        {
            foreach (Flag flag in _harvestingTruthTable.Rows.FirstOrDefault().Flags)
            {
                if (!flag.IsSum)
                {
                    continue;
                }

                int enableFlagCount = flag.AllFlags.Count(enableFlags.Contains);

                foreach (string range in _harvestingTruthTable.MergeValueByFlag[flag.FlagName])
                {
                    int.TryParse(range.Replace("To", "-").Split('-').FirstOrDefault(), out int min);
                    int.TryParse(range.Replace("To", "-").Split('-').LastOrDefault(), out int max);
                    if (enableFlagCount >= min && enableFlagCount <= max)
                    {
                        enableFlags.Add($"{flag.FlagName}_{range}");
                    }
                }
            }
        }
        private class BistInfo
        {
            public string HarvsetFlag { get; set; }
            public List<string> Levels { get; set; } = new List<string>();
        }
    }
}
