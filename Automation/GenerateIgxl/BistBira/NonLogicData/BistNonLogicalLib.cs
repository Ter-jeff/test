using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BistBira.Base;
using Automation.GenerateIgxl.BistBira.NewLogicData;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using ScghLib.Enums;
using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.Harvest;
using TestPlanLib.VbtLib;

using BistAction = ScghLib.Utility.BistAction;

namespace Automation.GenerateIgxl.BistBira.NonLogicData
{
    public class BistNonLogicalLib
    {
        internal class GenInstanceRow
        {
            public string GetColumnA(BistProdFlowRow bistProdFlowSheetRow)
            {
                return bistProdFlowSheetRow.SheetName + ",Row" + bistProdFlowSheetRow.RowNum;
            }

            public string GetTimeSetName(string patternName)
            {
                Dictionary<string, PatternData> patternListDic = AcTSetCategoryMapSingleton.Instance().PatternList;
                if (patternListDic == null)
                {
                    return "";
                }

                string timeSet = "";
                if (patternListDic.TryGetValue(patternName.ToLower(), out PatternData targetPatternData))
                {
                    timeSet = targetPatternData.TimeSetVersion;
                }

                return timeSet;
            }

            public string GetTimeSetVersion(string timeSet)
            {
                Dictionary<string, int> timeSetDic = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
                if (timeSetDic == null)
                {
                    return "missing timeset";
                }

                int version = -1;
                if (timeSetDic.TryGetValue(timeSet, out int value))
                {
                    version = value;
                }


                return version == -1 ? "missing timeset" : timeSet + "_" + version;
            }
        }

        private BistProdFlowSheet _prodFlowSheet;
        private readonly BistNaming _bistNaming;

        private string _module;
        private const string ConVbt = "VBT";
        private const string ConCSharpFunctional = "FuncTestMain";
        private const string ConFunctionalTUpdated = "Functional_T_Updated";
        private const string ConCSharpHarvestSummary = "HarvestSummary";
        private const string ConMbistRscr = "Mbist_RSCR";
        private const string ConCSharpMbistRetentionLevelWait = "MbistRetentionLevelWaitMain";
        private const string ConMbistRetentionLevelWait = "MbistRetentionLevelWait";
        private const string ConMbistRboxselDigCapinfo = "Mbist_RBOX_SEL_DigCap_info";
        private const string ConMbist = "Mbist";
        private const string ConTyp = "Typ";
        private const string ConMax = "Max";
        private const string ConMin = "Min";
        private const string ConLevels = "Levels";
        private const string ConPercore = @"PERCORE\d*";
        public const string ConIrt = "IRT";
        public const string ConErt = "ERT";
        public const string ConBir = "BIR";
        public const string ConErtbira = "ERTBIRA";
        public const string ConErtbist = "ERTBIST";
        public const string ConNrt = "NRT";
        public const string ConSrt = "SRT";
        public const string ConSrtbira = "SRTBIRA";
        public const string ConWus = "WUS";
        public const string ConEfc = "EFC";
        private string _comment = "";
        private readonly MultiTestSettingSheetsSingleton _multiTestSettings;
        private readonly BinCutFlowTable _equationVoltage;
        public Dictionary<string, List<BistProdFlowRow>> LabelDic;

        public BistNonLogicalLib(MbistConfig pConfig, MultiTestSettingSheetsSingleton multiTestSettings, BinCutFlowTable equationVoltage = null)
        {
            _bistNaming = new BistNaming(pConfig);
            _multiTestSettings = multiTestSettings;
            _equationVoltage = equationVoltage;
        }

        /// The work flow to generate instance
        public void WorkFlow(BistProdFlowSheet pProdFlowSheet, BistIgxlResult pIgxlResult, MbistDataStore mbistDataStore, MbistPatSetType mbistPatSetType)
        {
            _prodFlowSheet = pProdFlowSheet;
            LabelDic = _prodFlowSheet.LabelDic;
            _module = _bistNaming.GetModule(_prodFlowSheet.MbistSheet.SheetName);
            string sheetName = "TestInst_" + _prodFlowSheet.MbistSheet.SheetName;

            if (Regex.IsMatch(pProdFlowSheet.MbistSheet.SheetName, ConPercore, RegexOptions.IgnoreCase))
            {
                _comment = Regex.Match(pProdFlowSheet.MbistSheet.SheetName, ConPercore, RegexOptions.IgnoreCase).ToString();
            }

            if (!pIgxlResult.InstanceSheets.Any(p => p.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                var instanceSheet = new InstanceSheet(sheetName);
                instanceSheet.SourceInfo.Block = nameof(EnumBlock.Mbist);
                pIgxlResult.InstanceSheets.Add(instanceSheet);
            }
            InstanceSheet targetInstanceSheet = pIgxlResult.InstanceSheets.Find(p => p.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            var instanceRows = new List<InstanceRow>();
            var dummyInstanceRows = new List<InstanceRow>();
            for (int i = 0; i < _prodFlowSheet.Rows.Count; i++)
            {
                BistProdFlowRow dataRow = _prodFlowSheet.Rows[i];
                if (BistAction.GetActionType(dataRow) == BistActionType.RunPattern)
                {
                    instanceRows.Add(AddInstItem(dataRow, mbistPatSetType, mbistDataStore));
                }
                if (BistAction.GetActionType(dataRow) == BistActionType.Retention)
                {
                    bool isWaitTimeOnly = false;
                    string rampStep = "5";
                    string retentionPins = "";
                    BistProdFlowRow vddRow = null;
                    if (i != 0)
                    {
                        if (BistAction.GetActionType(_prodFlowSheet.Rows[i - 1]) != BistActionType.RetentionVoltDrop)
                        {
                            isWaitTimeOnly = true;
                        }
                        else
                        {
                            vddRow = _prodFlowSheet.Rows[i - 1];
                            rampStep = BistAction.GetRetentionRampStep(_prodFlowSheet.Rows[i - 1]);
                            retentionPins = _prodFlowSheet.Rows[i - 1].Note.Contains("RET_Pins") ? _prodFlowSheet.Rows[i - 1].Note.Split(':').Last() : "";
                        }
                    }

                    if (vddRow != null)
                    {
                        if (string.IsNullOrEmpty(vddRow.Pattern) ^ string.IsNullOrEmpty(dataRow.Pattern))
                        {
                            if (string.IsNullOrEmpty(vddRow.Pattern))
                            {
                                Response.Report($"Argument \"patternBeforeWait\" for retention will be empty because pattern column in row {vddRow.RowNum} is missing.", EnumMessageLevel.Error);
                                ErrorReportManager.AddError(MbistErrorType.E_BurstInfoMismatch_04, vddRow.SheetName, vddRow.RowNum, 1, "Argument \"patternBeforeWait\" for retention will be empty.");
                            }
                            if (string.IsNullOrEmpty(dataRow.Pattern))
                            {
                                Response.Report($"Argument \"patternAfterWait\" for retention will be empty because pattern column in row {dataRow.RowNum} is missing.", EnumMessageLevel.Error);
                                ErrorReportManager.AddError(MbistErrorType.E_BurstInfoMismatch_04, dataRow.SheetName, dataRow.RowNum, 1, "Argument \"patternAfterWait\" for retention will be empty.");
                            }
                        }
                    }

                    AddRetentionInstItem(instanceRows, dataRow, isWaitTimeOnly ? "" : vddRow?.Pattern, isWaitTimeOnly, rampStep, retentionPins);
                }
                if (BistAction.GetActionType(dataRow) == BistActionType.Null && dataRow.PassBranch != dataRow.FailBranch)
                {
                    AddDummyInstItem(dummyInstanceRows, dataRow);
                }
            }

            AddHarvestSummary(sheetName, instanceRows);
            var existInst = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (InstanceRow instanceRow in instanceRows)
            {
                if (existInst.Add(instanceRow.TestName))
                {
                    targetInstanceSheet.AddRow(instanceRow);
                }
            }

            foreach (InstanceRow dummyInst in dummyInstanceRows)
            {
                pIgxlResult.AddDummyInstance(sheetName, dummyInst);
            }
        }

        private void AddHarvestSummary(string sheetName, List<InstanceRow> pInstanceRows)
        {
            if (TestPlanStatic.HarvestingTruthTableSheets.Any())
            {
                HarvestingTruthTableSheet truthTable = TestPlanStatic.HarvestingTruthTableSheets.First();
                HarvestingTruthTableRow row1 = truthTable.Rows.First();
                List<string> allFlags = truthTable.GetHarvAllFlagList();
                if (TestPlanStatic.FlagOperationSheets.Any())
                {
                    foreach (FlagOperationSheet sheet in TestPlanStatic.FlagOperationSheets)
                    {
                        foreach (string flag in sheet.GetAllFlags())
                        {
                            if (!allFlags.Exists(x => x.Equals(flag, StringComparison.OrdinalIgnoreCase)))
                            {
                                allFlags.Add(flag);
                            }
                        }
                    }
                }
                var summaryInstance = new InstanceRow { TestName = BistNaming.GetHarvestInstName(sheetName.Replace("TestInst_", "")) };
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCSharpHarvestSummary, "mbist", true);
                if (function.Type == ".NET")
                {
                    summaryInstance.VbtType = ".NET";
                    summaryInstance.VbtName = function.FullFunctionName;
                    function.SetParamValue("allFlags", string.Join("+", allFlags.Distinct().Where(x => x != "")));
                    function.SetParamValue("harvestingFlags", row1.GetHarvGlobalFailFlag());
                    function.SetParamValue("sumFailCoreFlags", truthTable.SumFailCoreFlags());
                    function.SetParamValue("binoutFlags", row1.HarvBinOutFailFlag());
                }
                else
                {
                    function = TestProgram.VbtFunctionLib.GetFunctionByName("Harvest_Summary", "mbist");
                    summaryInstance.VbtType = "VBT";
                    summaryInstance.VbtName = function.FullFunctionName;
                    function.SetParamValue("Harv_AllFlag", string.Join("+", allFlags.Distinct().Where(x => x != "")));
                    function.SetParamValue("Harv_GlobalFailFlag", row1.GetHarvGlobalFailFlag());
                    function.SetParamValue("Harv_FailCoreSumFlag", truthTable.HarvFailCoreSumFlag());
                    function.SetParamValue("Harv_AllCorePassFlag", row1.HarvAllCorePassFlag());
                    function.SetParamValue("Harv_BinOutFailFlag", row1.HarvBinOutFailFlag());
                    function.SetParamValue("CustHarv_GlobalFailFlag", row1.GetCustHarvGlobalFailFlag());
                    function.SetParamValue("CustHarv_FailCoreSumFlag", truthTable.CustHarvFailCoreSumFlag());
                    function.SetParamValue("CustHarv_BinOutFailFlag", row1.CustHarvBinOutFailFlag());
                }
                summaryInstance.ArgList = function.Parameters;
                summaryInstance.Args = function.ArgList;
                pInstanceRows.Add(summaryInstance);
            }
        }

        private void AddRetentionInstItem(List<InstanceRow> pInstanceRows, BistProdFlowRow pRow = null, string vddRowPattern = "", bool isWaitTimeOnly = false, string rampStep = "5", string retentionPins = "")
        {
            var row = new InstanceRow();
            string lStrModule = _module;
            string chiplet = pRow?.Chiplet;
            string waitTime = BistAction.GetRetentionTime(pRow);
            string sheetName = _prodFlowSheet.MbistSheet.IsMultipleSheet ? _prodFlowSheet.MbistSheet.SheetName : "";
            string findRetentionInstanceRowTestName = FindRetentionInstanceTestName(pInstanceRows, pRow);

            string lstInstanceType = CheckRetentionNew(findRetentionInstanceRowTestName);

            row.TestName = _bistNaming.CreateRetentionTestNameNew(_module, pRow?.Voltage.Split(',', ' ').First(), sheetName, waitTime, rampStep, _bistNaming.GetWaitTimeOnly(isWaitTimeOnly));
            SetRetentionFunctionArgs(row, pRow, vddRowPattern, isWaitTimeOnly, rampStep, retentionPins, waitTime);

            var patterns = new List<string> { "Vdip" };


            SetRetentionDcCategory(row, pRow, lstInstanceType, lStrModule, chiplet, patterns);

            if (string.IsNullOrEmpty(row.DcSelector))
            {
                row.DcSelector = "Min";
            }

            string timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(pInstanceRows.Last().TimeSets, BlockType.Mbist);
            if (timeSet2Cat == "TBD")
            {
                timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(pInstanceRows.Last().TimeSets);
            }

            row.AcCategory = timeSet2Cat;
            row.AcSelector = ConTyp;

            row.TimeSets = pInstanceRows.Last().TimeSets;
            row.PinLevels = ConLevels + "_" + ConMbist;
            if (!pInstanceRows.Exists(x => x.TestName.Equals(row.TestName)))
            {
                pInstanceRows.Add(row);
            }
        }

        private string FindRetentionInstanceTestName(List<InstanceRow> pInstanceRows, BistProdFlowRow pRow)
        {
            string findRetentionInstanceRowTestName = "";

            for (int i = pInstanceRows.Count - 1; i >= 0; i--)
            {
                if (_bistNaming.GetPatternType(pInstanceRows[i].PayloadList.First()) == BistConst.ConPayload)
                {
                    findRetentionInstanceRowTestName = pInstanceRows[i].TestName;
                    break;
                }
            }

            if (!string.IsNullOrEmpty(pRow?.Voltage))
            {
                findRetentionInstanceRowTestName = pRow.Voltage;
            }

            return findRetentionInstanceRowTestName;
        }

        private void SetRetentionFunctionArgs(InstanceRow row, BistProdFlowRow pRow, string vddRowPattern, bool isWaitTimeOnly, string rampStep, string retentionPins, string waitTime)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCSharpMbistRetentionLevelWait, "mbist", true);
            if (function.Type == ".NET")
            {
                row.VbtType = ".NET";
                row.VbtName = function.FullFunctionName;
                row.ArgList = function.Parameters;
                row.Args = function.ArgList;
                #region SetArg
                function.SetParamValue("waitTimeMs", waitTime);
                function.SetParamValue("rampSteps", rampStep);
                function.SetParamValue("stepWaitTime", "0.001");
                function.SetParamValue("waitTimeOnly", isWaitTimeOnly ? "1" : "0");
                function.SetParamValue("rampPins", string.IsNullOrEmpty(retentionPins) ? "" : retentionPins);
                function.SetParamValue("bypassDcSpec", GetBypassDcSpecValue(MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.IsSplitDcSpecs(LocalSpecs.Options.IsSplitDcSpecs)));
                if (!string.IsNullOrEmpty(vddRowPattern) && !string.IsNullOrEmpty(pRow?.Pattern))
                {
                    function.SetParamValue("patternsBeforeWait", vddRowPattern);
                    function.SetParamValue("patternsAfterWait", pRow.Pattern);
                }
                #endregion
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(ConMbistRetentionLevelWait, "mbist");
                row.VbtType = ConVbt;
                row.VbtName = function.FullFunctionName;
                row.ArgList = function.Parameters;
                row.Args = function.ArgList;
                #region SetArg
                function.SetParamValue("mS_Time", waitTime);
                function.SetParamValue("RampStep", rampStep);
                function.SetParamValue("RampWaitTime", "");
                function.SetParamValue("WaitTimeOnly", isWaitTimeOnly ? "1" : "0");
                function.SetParamValue("RET_TTR", rampStep == "0" ? "TRUE" : "FALSE");
                function.SetParamValue("RET_Pins", string.IsNullOrEmpty(retentionPins) ? "" : retentionPins);
                #endregion
            }
        }

        private void SetRetentionDcCategory(InstanceRow row, BistProdFlowRow pRow, string lstInstanceType, string lStrModule, string chiplet, List<string> patterns)
        {
            if (string.IsNullOrEmpty(lstInstanceType) && string.IsNullOrEmpty(pRow?.Voltage))
            {
                row.DcCategory = "";
            }
            else if (!string.IsNullOrEmpty(pRow?.Voltage))
            {
                string[] arr = pRow.Voltage.Split(',', ' ');
                string selector = arr.Length > 1 ? arr[1] : "";
                row.DcCategory = arr[0];
                if (selector != "")
                {
                    row.DcSelector = GetSelector(selector);
                }

                if (!_multiTestSettings.DcCategoryInfos.Exists(s => s.CategoryName.Equals(arr[0], StringComparison.OrdinalIgnoreCase)))
                {
                    row.DcCategory += "(Should have)";
                }
            }
            else
            {
                row.DcCategory = _multiTestSettings.FindMbistCatgeoryName(lStrModule, lstInstanceType, "Retention", patterns, out EnumMessageLevel msgLevel, out string errorMsg, chiplet);
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    Response.Report(errorMsg, msgLevel, 0);
                }
            }
        }

        internal static string GetBypassDcSpecValue(bool isSplitDc)
        {
            return isSplitDc ? "_SC" : string.Empty;
        }

        private string GetDigSource(BistProdFlowRow pSheetRow)
        {
            string lBistorBira = CheckBistOrBira(pSheetRow);
            var patterns = new List<string> { pSheetRow.Pattern };
            string payloadDcSpec;
            string payloadDcSelector;
            string chiplet = pSheetRow.Chiplet;
            string sendPin = "JTAG_TDI";
            string selsramPattern;
            if (pSheetRow.IsDsscRow)
            {
                selsramPattern = pSheetRow.Pattern;
            }
            else
            {
                selsramPattern = pSheetRow.BurstPatterns.FirstOrDefault(x => new BistNaming(new MbistConfig()).IsSelDssc(x));
            }
            if (!string.IsNullOrEmpty(selsramPattern))
            {
                string pModule = _bistNaming.GetPatternModule(pSheetRow.VoltageMode, _module);
                string payloadVoltageFromScgh = pSheetRow.Voltage.Split(',', ' ').First();
                if (string.IsNullOrEmpty(pSheetRow.Voltage) && !_multiTestSettings.DcCategoryInfos.Exists(s => s.CategoryName.Equals(payloadVoltageFromScgh, StringComparison.OrdinalIgnoreCase)))
                {
                    payloadDcSpec = _multiTestSettings.FindMbistCatgeoryName(_module, lBistorBira, pSheetRow.VoltageMode, patterns, out EnumMessageLevel _, out string _, chiplet, pModule);
                }
                else
                {
                    payloadDcSpec = pSheetRow.Voltage.Split(',', ' ').First();
                }
                payloadDcSelector = GetDcSelector(BistNaming.GetVoltageType(pSheetRow.Voltage));
                if (LocalSpecs.HardIpInfos != null)
                {
                    string foundPin = LocalSpecs.HardIpInfos.GetHardIpInfo(selsramPattern)?.SendPinName;
                    if (!string.IsNullOrEmpty(foundPin))
                    {
                        sendPin = foundPin;
                    }
                }
                if (payloadDcSelector.Equals("Typ", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Test_" + payloadDcSpec + "_NV:" + sendPin;
                }

                if (payloadDcSelector.Equals("Max", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Test_" + payloadDcSpec + "_HV:" + sendPin;
                }

                if (payloadDcSelector.Equals("Min", StringComparison.CurrentCultureIgnoreCase))
                {
                    return "Test_" + payloadDcSpec + "_LV:" + sendPin;
                }
            }
            return "";
        }

        /// Add a InstanceRow into list
        private InstanceRow AddInstItem(BistProdFlowRow bistProdFlowSheetRow, MbistPatSetType mbistPatSetType, MbistDataStore mbistDataStore)
        {
            var row = new InstanceRow();
            var genInstanceRow = new GenInstanceRow();
            row.ColumnA = genInstanceRow.GetColumnA(bistProdFlowSheetRow);
            string lStrModule = _module;
            string lSheetModule = _module;
            _bistNaming.GetPatternModule(bistProdFlowSheetRow.VoltageMode, lStrModule);

            if (bistProdFlowSheetRow.IsRscr())
            {
                GenRscrTestMethod(bistProdFlowSheetRow, row, mbistPatSetType);
            }
            else if (bistProdFlowSheetRow.Note.Contains("Func:"))
            {
                GenDefFuncTestMethod(bistProdFlowSheetRow, row);
            }
            else
            {
                GenTestMethod(bistProdFlowSheetRow, row, mbistPatSetType, mbistDataStore);
            }
            //Dc spec
            GetDcCatAndSelectorByNewTs(bistProdFlowSheetRow, row);
            bistProdFlowSheetRow.DcCategory = row.DcCategory;
            row.ColumnA += ";DC Category:" + row.DcCategory;
            row.TestName = _bistNaming.CreateNewTestName(lSheetModule, bistProdFlowSheetRow);

            if (!string.IsNullOrEmpty(bistProdFlowSheetRow.NextPayLoadPattern)) // means this is DSSC row 
            {
                row.PinLevels = ConLevels + "_" + ConMbist;
            }
            else
            {
                row.PinLevels = ConLevels + "_" + ConMbist;
                if (bistProdFlowSheetRow.Pattern.Contains("EFP") &&
                    bistProdFlowSheetRow.DcCategory.StartsWith("HardIP", StringComparison.OrdinalIgnoreCase))
                {
                    row.PinLevels = ConLevels + "_" + "HardIp";
                }
            }
            if (TestProgram.VbtFunctionLib.GetFunctionByName(ConCSharpFunctional, ConMbist, true).Type != ".NET" &&
                bistProdFlowSheetRow.Pattern.Contains("EFP"))
            {
                row.Overlay = _module + ConMbist + "_" + _module + "_RepairEfuse";
            }
            string timeSets = genInstanceRow.GetTimeSetName(bistProdFlowSheetRow.Pattern);
            string userDefineAc = "";
            if (!timeSets.ContainsIgnoreCase(bistProdFlowSheetRow.TimeSet.Split('.').First().ToUpper()))
            {
                string newTimeSet = bistProdFlowSheetRow.TimeSet.Split(':').First();
                timeSets = genInstanceRow.GetTimeSetVersion(newTimeSet.Split('.').First().ToUpper());

                if (bistProdFlowSheetRow.TimeSet.Contains(":"))
                {
                    userDefineAc = bistProdFlowSheetRow.TimeSet.Split(':').Last();
                }
                else
                {
                    if (timeSets != "")
                    {
                        timeSets += "(Mismatch)";
                    }
                }
            }
            string timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets, BlockType.Mbist);
            if (timeSet2Cat == "TBD")
            {
                timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets);
            }

            row.AcCategory = string.IsNullOrEmpty(userDefineAc) ? timeSet2Cat : userDefineAc;
            row.AcSelector = ConTyp;
            row.TimeSets = timeSets;
            row.Comment = _comment;
            return row;
        }

        private void GenTestMethod(BistProdFlowRow bistProdFlowSheetRow, InstanceRow row, MbistPatSetType mbistPatSetType, MbistDataStore mbistDataStore)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCSharpFunctional, ConMbist, true);
            if (function.Type == ".NET")
            {
                GenerateCSharpInstanceRow(row, bistProdFlowSheetRow, function, mbistPatSetType);
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(ConFunctionalTUpdated, "mbist");
                GenerateVbtInstanceRow(row, bistProdFlowSheetRow, function, mbistDataStore);
            }
        }

        private void GenerateCSharpInstanceRow(InstanceRow row, BistProdFlowRow bistProdFlowSheetRow, Function function, MbistPatSetType mbistPatSetType)
        {
            row.VbtType = ".NET";
            row.VbtName = function.FullFunctionName;
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;
            row.PayloadList.Add(bistProdFlowSheetRow.Pattern);
            if (bistProdFlowSheetRow.Voltage.Contains("_EQN") && LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                int idx = function.Parameters.Split(',').ToList().FindIndex(x => x.Equals("ateTestCondition"));
                if (idx != -1)
                {
                    row.Args[idx] = bistProdFlowSheetRow.Voltage;
                }
            }

            var dsscPat = new List<string>();
            var digSrcEquation = new List<string>();
            var digSrcAssignment = new List<string>();
            List<string> patList = bistProdFlowSheetRow.BurstPatterns.Any() ? bistProdFlowSheetRow.BurstPatterns : new List<string> { bistProdFlowSheetRow.Pattern };
            List<string> dsscPatterns = TestPlanStatic.UfDigSrcSheets != null && TestPlanStatic.UfDigSrcSheets.Any() ? TestPlanStatic.UfDigSrcSheets.SelectMany(x => x.Rows).Where(y => !string.IsNullOrEmpty(y.PatternName)).Select(z => z.PatternName).ToList() : new List<string>();

            foreach (string pat in patList)
            {
                if (dsscPatterns.Exists(x => x.Equals(pat, StringComparison.CurrentCultureIgnoreCase)))
                {
                    dsscPat.Add(pat);
                    digSrcEquation.Add("D");
                }
                else if (_bistNaming.IsSelDssc(pat))
                {
                    dsscPat.Add(pat);
                    digSrcEquation.Add("C");
                }
                else
                {
                    digSrcEquation.Add("");
                }
            }

            if (dsscPat.Any())
            {
                string sendPinName = "JTAG_TDI";
                string patternName = dsscPat.FirstOrDefault();
                if (LocalSpecs.HardIpInfos != null)
                {
                    HardIpInfo target = LocalSpecs.HardIpInfos.GetHardIpInfo(patternName);
                    sendPinName = string.IsNullOrEmpty(target.SendPinName) ? sendPinName : target.SendPinName;
                }
                function.SetParamValue("digSrcPin", sendPinName);
                function.SetParamValue("digSrcEquation", string.Join("|", digSrcEquation));
                if (digSrcEquation.Contains("C"))
                {
                    string syntax = "";
                    bool isEqn = false;
                    GetSelsrmSyntaxItself(bistProdFlowSheetRow, ref syntax);
                    if (string.IsNullOrEmpty(syntax))
                    {
                        GetSelsrmSyntax(bistProdFlowSheetRow, ref syntax, ref isEqn, mbistPatSetType);
                    }

                    if (string.IsNullOrEmpty(syntax) && !isEqn)
                    {
                        ErrorReportManager.AddError(MbistErrorType.E_Business_02, _prodFlowSheet.MbistSheet.SheetName, bistProdFlowSheetRow.RowNum, 1, $"Cannot find any payload from pass branch : {bistProdFlowSheetRow.Label} , selsram will use init pattern category.", new string[] { bistProdFlowSheetRow.Label });
                    }
                    digSrcAssignment.Add($"C=Selsram({syntax})");
                }
                if (digSrcEquation.Contains("D"))
                {
                    digSrcAssignment.Add("D=DSSC()");
                }
                function.SetParamValue("digSrcAssignment", string.Join(";", digSrcAssignment));
            }
            function.SetParamValue("Patterns", bistProdFlowSheetRow.Pattern);
            function.SetParamValue("ResultMode", mbistPatSetType == MbistPatSetType.BurstYes ? "1" : "0");
            function.SetParamValue("RelayMode", "1");
        }

        private void GetSelsrmSyntaxItself(BistProdFlowRow bistProdFlowSheetRow, ref string syntax)
        {
            bool isBurst = bistProdFlowSheetRow.BurstPatterns.Any();
            bool foundPl = false;
            string currentLabel = bistProdFlowSheetRow.Label;
            IEnumerable<string> sameLabelItems = isBurst ? bistProdFlowSheetRow.BurstPatterns : LabelDic[currentLabel].Select(x => x.Pattern);
            bool foundPat = false;
            foreach (string pattern in sameLabelItems)
            {
                if (_bistNaming.IsSelDssc(pattern))
                {
                    foundPat = true;
                    continue;
                }
                if (foundPat && _bistNaming.GetPatternType(pattern) == BistConst.ConPayload)
                {
                    foundPl = true;
                }
            }
            if (foundPl)
            {
                string voltage = !string.IsNullOrEmpty(bistProdFlowSheetRow.EqnVoltage) ? bistProdFlowSheetRow.EqnVoltage : bistProdFlowSheetRow.Voltage;
                string dcCategory = FindDcCategory(voltage, "");
                string selector = FindDcSelector(voltage);
                if (!string.IsNullOrEmpty(dcCategory) && !string.IsNullOrEmpty(selector))
                {
                    syntax = $"#{dcCategory}#{selector}";
                }
            }
        }

        private void GetSelsrmSyntax(BistProdFlowRow bistProdFlowSheetRow, ref string syntax, ref bool isEqn, MbistPatSetType mbistPatSetType)
        {
            string passBranch = LabelDic[bistProdFlowSheetRow.Label].Last().PassBranch;
            List<BistProdFlowRow> targetBranches = LabelDic[passBranch];
            BistProdFlowRow firstPayloadBranch = targetBranches.Find(x => mbistPatSetType != MbistPatSetType.Single ? x.BurstPatterns.Exists(y => _bistNaming.GetPatternType(y) == BistConst.ConPayload) : _bistNaming.GetPatternType(x.Pattern) == BistConst.ConPayload);
            if (firstPayloadBranch == null)
            {
                if (targetBranches.Any())
                {
                    GetSelsrmSyntax(targetBranches.First(), ref syntax, ref isEqn, mbistPatSetType);
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(firstPayloadBranch.EqnVoltage))
                {
                    isEqn = true;
                    syntax = "";
                    return;
                }
                string voltage = !string.IsNullOrEmpty(firstPayloadBranch.EqnVoltage) ? firstPayloadBranch.EqnVoltage : firstPayloadBranch.Voltage;
                string dcCategory = FindDcCategory(voltage, "");
                string selector = FindDcSelector(voltage);
                if (!string.IsNullOrEmpty(dcCategory) && !string.IsNullOrEmpty(selector))
                {
                    syntax = $"#{dcCategory}#{selector}";
                }
            }
        }

        private void GenerateVbtInstanceRow(InstanceRow row, BistProdFlowRow bistProdFlowSheetRow, Function vbt, MbistDataStore mbistDataStore)
        {
            vbt.SetParamValue("DigSource", GetDigSource(bistProdFlowSheetRow));
            row.VbtType = ConVbt;
            row.VbtName = vbt.FullFunctionName;
            row.ArgList = vbt.Parameters;
            row.Args = vbt.ArgList;
            row.Args[0] = bistProdFlowSheetRow.Pattern;
            row.PayloadList.Add(bistProdFlowSheetRow.Pattern);
            row.Args[8] = "0";
            row.Args[24] = "1";
            if (bistProdFlowSheetRow.Voltage.Contains("_EQN") && LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                int idx = vbt.Parameters.Split(',').ToList().FindIndex(x => x.Equals("ApplyVoltageFromBinCut"));
                if (idx != -1)
                {
                    row.Args[idx] = bistProdFlowSheetRow.Voltage;
                }
            }
            string sheetName = bistProdFlowSheetRow.SheetName;
            //Enable Finger print
            List<PatSet> patSetList = mbistDataStore.DicPatSets.TryGetValue(sheetName, out List<PatSet> set) ? set : null;
            if (LocalSpecs.FingerPrintList != null && LocalSpecs.FingerPrintList.Exists(p => p.Equals(bistProdFlowSheetRow.Pattern, StringComparison.OrdinalIgnoreCase)))
            {
                row.Args[36] = "-1";
            }
            else if (LocalSpecs.FingerPrintList != null && patSetList != null && patSetList.Exists(p => p.PatSetName.Equals(bistProdFlowSheetRow.Pattern)))
            {
                List<PatSetRow> patSetRows = patSetList.Find(p => p.PatSetName.Equals(bistProdFlowSheetRow.Pattern)).PatSetRows;
                foreach (PatSetRow psRow in patSetRows)
                {
                    if (LocalSpecs.FingerPrintList.Exists(p => p.Equals(psRow.File, StringComparison.OrdinalIgnoreCase)))
                    {
                        row.Args[36] = "-1";
                        break;
                    }
                }
            }

            //For Mbist powerup, poweroff efuse item
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                foreach (string notestr in bistProdFlowSheetRow.Note.Split(';'))
                {
                    if (notestr.StartsWith("rbox_program:", StringComparison.CurrentCultureIgnoreCase) ||
                        notestr.StartsWith("sleep_assert:", StringComparison.CurrentCultureIgnoreCase))
                    {
                        string rampinfo = notestr.Split(':')[1].Trim();
                        vbt.SetParamValue("StartOfBodyFArgs", rampinfo.Split('|')[0]);
                        vbt.SetParamValue("EndOfBodyFArgs", rampinfo.Split('|')[1]);
                        break;
                    }
                }
            }
        }

        private void GenRscrTestMethod(BistProdFlowRow bistProdFlowSheetRow, InstanceRow row, MbistPatSetType mbistPatSetType)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCSharpFunctional, "mbist", true);
            if (function.Type == ".NET")
            {
                GenerateCSharpInstanceRow(row, bistProdFlowSheetRow, function, mbistPatSetType);
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(ConMbistRscr, "mbist");
                GenerateRscrVbtInstanceRow(row, bistProdFlowSheetRow, function);
            }
        }

        private void GenerateRscrVbtInstanceRow(InstanceRow row, BistProdFlowRow bistProdFlowSheetRow, Function vbt)
        {
            vbt.SetParamValue("Shift_Pat", bistProdFlowSheetRow.Pattern);
            string[] patItems = bistProdFlowSheetRow.Pattern.Split('_');
            foreach (string patItem in patItems)
            {
                if (Regex.IsMatch(patItem, @"SRV\d+", RegexOptions.IgnoreCase))
                {
                    vbt.SetParamValue("MBIST_BLOCK", _module + "_" + patItem);
                    vbt.SetParamValue("Server", patItem);
                }
            }
            row.VbtType = ConVbt;
            row.VbtName = vbt.FullFunctionName;
            row.ArgList = vbt.Parameters;
            row.Args = vbt.ArgList;
        }

        private void GenDefFuncTestMethod(BistProdFlowRow bistProdFlowSheetRow, InstanceRow row)
        {
            string deffuncname = "";
            foreach (string noteinfo in bistProdFlowSheetRow.Note.Split(';'))
            {
                if (noteinfo.Equals("Func:" + ConMbistRboxselDigCapinfo, StringComparison.CurrentCultureIgnoreCase))
                {
                    deffuncname = ConMbistRboxselDigCapinfo;
                }
            }

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCSharpFunctional, "mbist", true);
            if (function.Type == ".NET")
            {

            }
            else if (deffuncname != "")
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(deffuncname, "mbist");
                GenerateDefFuncVbtInstanceRow(row, bistProdFlowSheetRow, function);
            }
        }

        private void GenerateDefFuncVbtInstanceRow(InstanceRow row, BistProdFlowRow bistProdFlowSheetRow, Function vbt)
        {
            vbt.SetParamValue("Pat", bistProdFlowSheetRow.Pattern);

            foreach (string arginfo in bistProdFlowSheetRow.Note.Split(';'))
            {
                if (!arginfo.Contains("="))
                {
                    continue;
                }

                string argKey = arginfo.Split('=')[0].Trim();
                string argValue = arginfo.Split('=')[1].Trim();
                vbt.SetParamValue(argKey, argValue);
            }

            row.VbtType = ConVbt;
            row.VbtName = vbt.FullFunctionName;
            row.ArgList = vbt.Parameters;
            row.Args = vbt.ArgList;
        }

        private void AddDummyInstItem(List<InstanceRow> pInstanceRows, BistProdFlowRow pRow)
        {
            var row = new InstanceRow();
            string lStrModule = _module;
            row.TestName = _bistNaming.CreateDummyTestName(lStrModule, pRow);  //2017/05/04 anderson add for MBist rename inst name
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ConCSharpFunctional, "mbist", true);
            if (function.Type == ".NET")
            {
                row.VbtType = ".NET";
                row.VbtName = function.FullFunctionName;
                row.ArgList = function.Parameters;
                row.Args = function.ArgList;
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(ConFunctionalTUpdated, "mbist");
                row.VbtType = ConVbt;
                row.VbtName = function.FullFunctionName;
                row.ArgList = function.Parameters;
                row.Args = function.ArgList;
            }

            GetDcCatAndSelectorByNewTs(pRow, row);

            var genInstanceRow = new GenInstanceRow();
            string timeSets = genInstanceRow.GetTimeSetName(pRow.Pattern);
            string timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets, BlockType.Mbist);
            if (timeSet2Cat == "TBD")
            {
                timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets);
            }

            row.AcCategory = timeSet2Cat;
            row.AcSelector = ConTyp;
            row.TimeSets = timeSets;
            row.PinLevels = ConLevels + "_" + ConMbist;
            row.Comment = _comment;
            pInstanceRows.Add(row);
        }

        private string GetDcCatAndSelectorByNewTs(BistProdFlowRow bistProdFlowSheetRow, InstanceRow instanceRow)
        {
            string rowPerformanceMode = bistProdFlowSheetRow.IsDsscRow ? bistProdFlowSheetRow.OriPerformance : bistProdFlowSheetRow.VoltageMode;
            string module = _module;
            module = _bistNaming.GetPatternModule(rowPerformanceMode, module);
            string voltage = bistProdFlowSheetRow.IsDsscRow ? rowPerformanceMode : bistProdFlowSheetRow.Voltage;
            if (!string.IsNullOrEmpty(bistProdFlowSheetRow.EqnVoltage) && !bistProdFlowSheetRow.IsDsscRow)
            {
                voltage = bistProdFlowSheetRow.EqnVoltage;
            }
            //Find DC Category
            string dcCategory = FindDcCategory(voltage, module);
            if (string.IsNullOrEmpty(dcCategory))
            {
                ErrorReportManager.AddError(MbistErrorType.E_Business_03, _prodFlowSheet.MbistSheet.SheetName, bistProdFlowSheetRow.RowNum, 1, $"Error! Can not found DC Spec by {voltage} in {bistProdFlowSheetRow.Label}", new string[] { voltage, bistProdFlowSheetRow.Label });
                instanceRow.DcCategory = $"Error! Can not found DC Spec by {voltage}";
            }
            else
            {
                string dcSelector = FindDcSelector(voltage);
                instanceRow.DcCategory = dcCategory;
                instanceRow.DcSelector = dcSelector;
            }
            return module;
        }

        internal string FindDcSelector(string pVoltage)
        {
            string dcSelector = "";
            string[] voltageinfoList = pVoltage.Split(',', ' ');
            if (voltageinfoList.Length == 2)
            {
                dcSelector = GetSelector(voltageinfoList[1]);
            }
            else
            {
                List<string> tSelector = _multiTestSettings.GetDcCategoryVoltages(pVoltage);
                if (tSelector.Count == 1)
                {
                    dcSelector = GetSelector(tSelector[0]);
                }
                else
                {
                    string voltageType = BistNaming.GetVoltageType(pVoltage);
                    #region set selector
                    if (voltageType == BistConst.ConHv || voltageType == BistConst.ConMhv || voltageType == BistConst.ConMhv1 || voltageType == BistConst.ConMhv2 || voltageType == BistConst.ConMhv3)
                    {
                        dcSelector = ConMax;
                    }
                    else if (voltageType == BistConst.ConLv || voltageType == BistConst.ConMlv || voltageType == BistConst.ConMlv1 || voltageType == BistConst.ConMlv2 || voltageType == BistConst.ConMlv3)
                    {
                        dcSelector = ConMin;
                    }
                    else if (voltageType == BistConst.ConNv)
                    {
                        dcSelector = ConTyp;
                    }
                    #endregion
                }
            }
            return dcSelector;
        }

        private string FindDcCategory(string pVoltage, string module)
        {
            //Rule-2
            string[] voltageinfoList = pVoltage.Split(',', ' ');
            if (voltageinfoList[0].ContainsIgnoreCase("_EQN") && _equationVoltage != null)
            {
                voltageinfoList[0] = voltageinfoList[0].Replace("_EQN", "");
            }
            bool tFound = _multiTestSettings.DcCategoryInfos.Exists(s => s.CategoryName.Equals(voltageinfoList[0], StringComparison.OrdinalIgnoreCase));
            string dcCategory = "";
            if (tFound)
            {
                dcCategory = voltageinfoList[0];
            }
            else
            {
                //Rule-3
                List<string> voltages = pVoltage.Split(',', ' ').ToList();
                if (voltages.Count == 2)
                {
                    var keys = new List<string>();
                    foreach (string item in voltages)
                    {
                        if (item.Equals(BistConst.ConMhv, StringComparison.InvariantCultureIgnoreCase))
                        {
                            keys.Add(BistConst.ConVmargin1);
                        }
                        else if (item.Equals(BistConst.ConMlv, StringComparison.InvariantCultureIgnoreCase))
                        {
                            keys.Add(BistConst.ConVmargin3);
                        }
                        else
                        {
                            keys.Add(item);
                        }
                    }
                    keys.Add(module);
                    string findName = _multiTestSettings.DcCategoryInfos.ContainsCategoryName(keys);
                    dcCategory = findName;
                }
            }
            return dcCategory;
        }

        internal string GetSelector(string voltage)
        {
            string lSelector;
            if (voltage.Equals(BistConst.ConNv, StringComparison.InvariantCultureIgnoreCase))
            {
                lSelector = ConTyp;
            }
            else if (voltage.Equals(BistConst.ConHv, StringComparison.InvariantCultureIgnoreCase))
            {
                lSelector = ConMax;
            }
            else if (voltage.Equals(BistConst.ConLv, StringComparison.InvariantCultureIgnoreCase))
            {
                lSelector = ConMin;
            }
            else
            {
                lSelector = ConTyp;
            }
            return lSelector;
        }

        internal string GetDcSelector(string voltageType)
        {
            if (voltageType == BistConst.ConHv || voltageType == BistConst.ConMhv || voltageType == BistConst.ConMhv1 || voltageType == BistConst.ConMhv2 || voltageType == BistConst.ConMhv3)
            {
                return ConMax;
            }
            if (voltageType == BistConst.ConLv || voltageType == BistConst.ConMlv || voltageType == BistConst.ConMlv1 || voltageType == BistConst.ConMlv2 || voltageType == BistConst.ConMlv3)
            {
                return ConMin;
            }
            if (voltageType == BistConst.ConNv)
            {
                return ConTyp;
            }
            return "";
        }

        private string CheckBistOrBira(BistProdFlowRow pattern)
        {
            string patternName = string.IsNullOrEmpty(pattern.NextPayLoadPattern)
                ? pattern.Pattern
                : pattern.NextPayLoadPattern;
            if (patternName.Contains("BIR"))
            {
                return "Bira";
            }

            return "Bist";
        }

        public static string CheckRetentionNew(string patternName)
        {
            if ((patternName.Contains("_" + ConErt + "_") && patternName.Contains(ConBir)) || patternName.Contains(ConErtbira))
            {
                return ConErtbira;
            }

            if ((patternName.Contains("_" + ConErt + "_") && !patternName.Contains(ConBir)) || patternName.Contains(ConErtbist))
            {
                return ConErtbist;
            }

            if (patternName.Contains("_" + ConNrt + "_"))
            {
                return ConNrt;
            }

            if (patternName.Contains("_" + ConSrtbira + "_"))
            {
                return ConSrtbira;
            }

            if (patternName.Contains("_" + ConSrt + "_"))
            {
                return ConSrt;
            }

            if (patternName.Contains("_" + ConWus + "_"))
            {
                return ConWus;
            }

            if (patternName.Contains("_" + ConEfc + "_"))
            {
                return ConEfc;
            }

            if (patternName.Contains("_" + ConIrt + "_"))
            {
                return ConIrt;
            }

            return string.Empty;
        }
    }
}
