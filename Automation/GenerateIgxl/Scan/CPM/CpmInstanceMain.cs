using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.PostAction.TempMon.Enums;
using Automation.GenerateIgxl.Scan.Harvest;
using Automation.GenerateIgxl.Scan.Harvest.Flow;
using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Atpg;
using Automation.Utility.Atpg.Data;
using Automation.Utility.Basic;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Harvest;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Scan.CPM
{
    public class CpmInstanceMain : ScanNonBinCutInstanceMain
    {
        protected override string Module => "CPM";
        private const string BinTableCpm = "Bin_Table_CPM";
        private const string DevChar = "DevChar_CPM_Sheet";
        private readonly CharSheet _charSheet = new CharSheet(DevChar);
        private BinCutFinalInstanceRows _binCutFinalInstanceRows;
        private readonly List<string> _allFlags = new List<string>();
        private readonly int _dsscBit;
        private string _lastSetupName = "";
        private string _lastCharInst = "";
        private readonly List<string> _existSteps = new List<string>();

        public CpmInstanceMain(ScanConfig config) : base(config)
        {
            ExcelWorksheet sheet = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRM_Mapping_Table"] ?? EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Mapping_Table"];
            if (sheet != null)
            {
                _dsscBit = new SelsrmMappingSheetReader().ReadSheet(sheet).Rows.GroupBy(x => x.Stage).First().ToList().FindAll(x => x.Pattern.ContainsIgnoreCase("_SRMDSSC")).Count;
            }
        }

        private List<InstanceRow> GenCpmInstances(IEnumerable<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            var instanceRows = new InstanceRows();
            var groups = binCutFinalInstanceRows.GroupBy(x => x.GetBlockByFlowName()).ToList();
            var existFlowHeaderFooter = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                foreach (BinCutFinalInstanceRow row in group)
                {
                    if (row.BinCutInstanceRow == null)
                    {
                        continue;
                    }
                    string subFlowName = row.GetSubFlowName();
                    Dictionary<string, string> specialSettings = row.BinCutInstanceRow.GetSpecialSettings();
                    if (specialSettings == null)
                    {
                        instanceRows.Add(GenCpmNormalInstance(row));
                    }
                    else
                    {
                        if (specialSettings.ContainsKey("WriteEfuse"))
                        {
                            instanceRows.Add(GenCpmEfuseWrite(row));
                        }
                        else if (specialSettings.ContainsKey("Function"))
                        {
                            instanceRows.AddRange(GenCpmInstance(row));
                        }
                        else if (specialSettings.ContainsKey("Fail") || specialSettings.ContainsKey("Pass"))
                        {
                            instanceRows.Add(GenCpmNormalInstance(row));
                        }
                    }
                    if (!string.IsNullOrEmpty(row.BinCutInstanceRow.Char))
                    {
                        GenCharRow(row);
                    }
                    existFlowHeaderFooter.Add(subFlowName);
                }
            }
            existFlowHeaderFooter.ToList().ForEach(x => instanceRows.AddHeaderFooter(x.ToUpper()));
            return instanceRows;
        }

        private InstanceRow GenCpmNormalInstance(BinCutFinalInstanceRow row)
        {
            var instanceRow = new InstanceRow();
            string selector = row.GetVoltageType();
            bool isChar = !string.IsNullOrEmpty(row.BinCutInstanceRow.Char);
            Function function;
            instanceRow.TestName = row.PatSetName;
            instanceRow.TimeSets = row.GetTimeSetVersion(row.PatternList);
            instanceRow.DcCategory = row.GetDcCategory();
            instanceRow.DcSelector = GetDcSelector(selector);
            instanceRow.AcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow.AcSpec) ? GetAcCategory(row.BinCutInstanceRow, instanceRow.TimeSets) : row.BinCutInstanceRow.AcSpec;
            instanceRow.AcSelector = "Typ";
            instanceRow.PinLevels = GenerateLevel(instanceRow.DcCategory, row.Block, row.BinCutInstanceRow.Levels);
            instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
            instanceRow.SheetName = row.BinCutInstanceRow.SheetName;
            instanceRow.Overlay = row.BinCutInstanceRow.Overlay;
            if (isChar)
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameFuncTestCharMain, "scan", true);
                if (function.Type == ".NET")
                {
                    _lastCharInst = instanceRow.TestName;
                    GenerateCsharpInstanceRowChar(row, ref instanceRow, function);
                }
                else
                {
                    GenerateVbtInstanceRowChar(row, ref instanceRow, function);
                }
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameFunctionalT, "scan", true);
                if (function.Type == ".NET")
                {
                    AtpgService.GenerateCSharpInstanceRow(row, instanceRow, function);
                }
                else
                {
                    GenerateVbtInstanceRow(row, ref instanceRow, function);
                }
            }
            TempMonService.TrySetTempMon(LocalSpecs.TempMonDatas, row.BinCutInstanceRow.TempMon, instanceRow.TestName, EnumType.Instance);
            return instanceRow;
        }

        private void GenerateVbtInstanceRow(BinCutFinalInstanceRow row, ref InstanceRow instanceRow, Function function)
        {
            row.GetVoltageType();

            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
                instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
            }

            string dsscPat = "";
            if (row.PatternList.Any())
            {
                foreach (string pat in row.PatternList)
                {
                    if (Regex.IsMatch(pat, @"\w*DSSC\w*", RegexOptions.IgnoreCase))
                    {
                        dsscPat = pat;
                    }
                }
            }
            function.SetParamValue("Patterns", row.PatternList.Count == 1 ? row.PatternList[0] : row.PatSetName);
            function.SetParamValue("ResultMode", "0");
            function.SetParamValue("RelayMode", "1");

            if (!string.IsNullOrEmpty(dsscPat))
            {
                string sendPinName = "JTAG_TDI";
                if (LocalSpecs.HardIpInfos != null)
                {
                    HardIpInfo target = LocalSpecs.HardIpInfos.GetHardIpInfo(dsscPat);
                    if (target != null && !string.IsNullOrEmpty(target.SendPinName))
                    {
                        sendPinName = target.SendPinName;
                    }
                }
                function.SetParamValue("DigSource", "Test_AutoSwitch:" + sendPinName.ToUpper());
            }


            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.VbtType = "VBT";
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
        }

        private void GenerateVbtInstanceRowChar(BinCutFinalInstanceRow row, ref InstanceRow instanceRow, Function function)
        {
            string selector = row.GetVoltageType();

            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
                instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
            }

            if (row.PatternList.Any())
            {
                if (row.InitList.Any())
                {
                    function.SetParamValue("INIT_Patset", row.InitPatSetName);
                }
                if (row.PayloadList.Any())
                {
                    function.SetParamValue("PL_Patset", row.PatSetName);
                }
                string paramValue = $"init_{selector}_pl_Sweep";
                function.SetParamValue("powerRunScenario", paramValue);
                if (row.PatternList.Exists(x => Regex.IsMatch(x, @"\w*_SRMDSSC\w*", RegexOptions.IgnoreCase)))
                {
                    function.SetParamValue("SELSRAM_DSSC", "SelSrm" + "".PadRight(_dsscBit, 'S'));
                }

                function.SetParamValue("pmode", row.GetDcCategory() + ":" + selector);
                function.SetParamValue("One_Time_INIT", "FALSE");
            }
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.VbtType = "VBT";
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
        }

        private void GenerateCsharpInstanceRowChar(BinCutFinalInstanceRow row, ref InstanceRow instanceRow, Function function)
        {
            string selector = row.GetVoltageType();
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
                instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
            }

            if (row.PatternList.Any())
            {
                if (row.InitList.Any())
                {
                    function.SetParamValue("initPatset", row.InitPatSetName);
                }
                if (row.PayloadList.Any())
                {
                    function.SetParamValue("plpatset", row.PatSetName);
                }
                string paramValue = $"init_{selector}_pl_Sweep";
                function.SetParamValue("powerRunScenario", paramValue);
                function.SetParamValue("pmode", row.GetDcCategory() + ":" + selector);
                function.SetParamValue("oneTimeInit", "FALSE");
            }

            function.SetParamValue("isHarvesting", row.BinCutInstanceRow?.IsHarvesting);
            function.SetParamValue("isPinMask", row.BinCutInstanceRow.EnableCoreMask);

            UserFunctionTableRow ufFuncSetting = null;
            if (!string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction) && TestPlanStatic.UserFunctionSheet != null)
            {
                ufFuncSetting = TestPlanStatic.UserFunctionSheet.Rows
                .FirstOrDefault(x => x.UserFunction.Equals(row.BinCutInstanceRow.UserFunction, StringComparison.OrdinalIgnoreCase));
            }
            if (ufFuncSetting != null)
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, function);
            }
            List<string> ufDigSrcPats = TestPlanStatic.UfDigSrcSheets
                .SelectMany(x => x.Rows)
                .Where(y => !string.IsNullOrEmpty(y.PatternName)).Select(z => z.PatternName).ToList();
            AtpgService.SetDigSrc(row, ufDigSrcPats, ufFuncSetting, LocalSpecs.HardIpInfos, "", row.PatternList, ref function);
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.VbtType = ".NET";
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
        }

        private void GenCharRow(BinCutFinalInstanceRow row)
        {
            string charSetting = row.BinCutInstanceRow.Char;

            var charSetup = new CharSetup();
            string setupName = charSetting.Split('=').First().Trim();
            _lastSetupName = setupName;
            List<string> settings = Regex.Match(charSetting.Split('=').Last(), @"\((?<setting>.+)\)", RegexOptions.IgnoreCase).Groups["setting"].ToString().Split('|').ToList();

            charSetup.SetupName = setupName;
            charSetup.TestMethod = CharSetupConst.TestMethodRetest;
            foreach (string setting in settings)
            {
                string pinName = setting.Split(':').First();
                string realPinName = string.Join("_", pinName.Split('_').ToList().FindAll(x => !x.Equals("VOP") && !x.Equals("VAR")));
                var charStep = new CharStep(setupName, pinName);
                string[] values = setting.Split(':').Last().Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                string shmooPinGlb = "Shmoo_" + pinName + "_GLB";
                List<GlobalSpec> specs = TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value?.Rows;
                if (specs != null && !specs.Any(s => s.Symbol.Equals(shmooPinGlb, StringComparison.OrdinalIgnoreCase)))
                {
                    TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(new GlobalSpec(shmooPinGlb, "0.087"));
                }

                charStep.Mode = "X Shmoo";
                charStep.ParameterType = "Global Spec";
                charStep.ParameterName = shmooPinGlb;
                charStep.RangeFrom = values[0];
                charStep.RangeTo = values[1];
                charStep.RangeStepSize = values[2];
                charStep.RangeCalcField = "Steps";
                charStep.AlgorithmName = values.Length == 4 ? values[3] : "Jump";
                charStep.AlgorithmArguments = "6";
                charStep.ApplyToPins = realPinName;
                charStep.ApplyToPinExecMode = "Simultaneous";
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharFuncNamePrintShmooInfoMain, "scan", true);
                if (function.Type == ".NET")
                {
                    charStep.PostStep = function.FullFunctionName;
                }
                else
                {
                    charStep.PostStep = "PrintShmooInfo";
                }

                charStep.PostStepArguments = "CorePower," + realPinName;
                charStep.OutputFormat = "Enhanced";
                charStep.OutputTextFile = "Disable";
                charStep.OutputSheet = "Disable";
                charStep.OutputSuspendDatalog = "TRUE";
                charStep.OutputDestinationsDatalog = "Enable";
                charStep.OutputDestinationsImmediateWin = "Disable";
                charStep.OutputDestinationsOutputWin = "Disable";
                if (_existSteps.Contains(setupName + ":" + pinName))
                {
                    continue;
                }

                _existSteps.Add(setupName + ":" + pinName);
                charSetup.CharSteps.Add(charStep);
            }
            if (charSetup.CharSteps.Any())
            {
                _charSheet.AddRow(charSetup);
            }
        }

        private InstanceRows GenCpmInstance(BinCutFinalInstanceRow row)
        {
            var instanceRows = new InstanceRows();
            var instanceRow = new InstanceRow();
            Dictionary<string, string> specialSettings = row.BinCutInstanceRow.GetSpecialSettings();
            string functionName = DcContiConst.VbtCreateOverlayCpm;
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
            }

            instanceRow.TestName = row.BinCutInstanceRow.PatSetNameOrange;

            instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
            instanceRow.SheetName = row.BinCutInstanceRow.SheetName;
            if (specialSettings.TryGetValue("Function", out string setting))
            {
                functionName = setting;
            }

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(functionName, "scan", true);
            foreach (KeyValuePair<string, string> specialSetting in specialSettings)
            {
                if (!specialSetting.Key.Equals("Function", StringComparison.CurrentCultureIgnoreCase))
                {
                    function.SetParamValue(specialSetting.Key, specialSetting.Value);
                }
            }
            function.SetParamValue("shmooStoredDataKey", _lastCharInst + "_" + _lastSetupName);

            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction))
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, function);
            }
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.VbtType = function.Type;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            instanceRows.Add(instanceRow);

            return instanceRows;
        }

        private InstanceRow GenCpmEfuseWrite(BinCutFinalInstanceRow row)
        {
            var instanceRow = new InstanceRow();
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpCpmEFuseWrite, "scan", true) ??
                                TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtCpmEFuseWrite, "scan");

            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
            }

            if (row.BinCutInstanceRow != null)
            {
                instanceRow.TestName = row.BinCutInstanceRow.PatSetNameOrange;
                instanceRow.VbtType = function.Type;
                instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
                instanceRow.SheetName = row.BinCutInstanceRow.SheetName;
                if (!string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction))
                {
                    TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, function);
                }
            }
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        protected void GenCpmInstanceSheet(List<InstanceRow> instanceRows)
        {
            if (instanceRows.Count > 0)
            {
                var instanceSheet = new InstanceSheet("TestInst_CPM");
                instanceSheet.AddRows(instanceRows);
                instanceSheet.RemoveDuplicateInstance(false);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCpm, instanceSheet);
            }
        }

        private List<BinCutInstanceSheet> GetCpmInstanceSheets()
        {
            var cpmInstanceSheets = new List<BinCutInstanceSheet>();
            _ = AcTSetCategoryMapSingleton.Instance().PatternList;
            _ = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
            foreach (BinCutInstanceSheet cpmInstanceSheet in TestPlanStatic.CpmInstanceSheets)
            {
                if (cpmInstanceSheet != null)
                {
                    cpmInstanceSheets.Add(cpmInstanceSheet);
                }
            }
            return cpmInstanceSheets;
        }

        public override void WorkFlow()
        {
            var patSets = new List<PatSet>();
            var flowSheets = new List<SubFlowSheet>();
            var instanceRows = new InstanceRows();
            var cpmSubFlows = new List<SubFlowSheet>();
            BinCutInstanceNamingSheet binCutInstanceNamingSheet = BinCutInstanceNamingSheet();
            List<BinCutInstanceSheet> cpmInstanceSheets = GetCpmInstanceSheets();
            var instSheetPreProcess = new InstSheetPreProcess(Config);
            BinCutFinalInstanceRows tpInstRows = instSheetPreProcess.InitialInstance(cpmInstanceSheets, binCutInstanceNamingSheet);
            _binCutFinalInstanceRows = tpInstRows.RePatSetNameDuplicateRows();
            AcSpecSheet acSpecSheet = TestProgram.IgxlWorkBk.GetAcSpecsSheet();
            if (acSpecSheet != null)
            {
                new BinCutAcSpecsWriter().GenAcSpecs(_binCutFinalInstanceRows, acSpecSheet);
            }

            //Gen CPM Instance
            instanceRows.AddRange(GenCpmInstances(_binCutFinalInstanceRows));

            // Gen CPM Flow
            cpmSubFlows.AddRange(GenFlow(tpInstRows));
            // Gen CPM MainFlow
            flowSheets.Add(GenCpmMainFlow(cpmSubFlows, ref instanceRows));
            flowSheets.AddRange(cpmSubFlows);
            // Gen BinTable
            BinTableSheet binTableSheet = GenCpmBinTable(_binCutFinalInstanceRows);
            // Gen PatSet
            patSets.AddRange(GenPatSets(tpInstRows));

            #region add into igxl
            TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirCpm, binTableSheet);
            SetPatSetSheet(patSets);
            foreach (SubFlowSheet flow in flowSheets)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirCpm, flow);
            }

            GenCpmInstanceSheet(instanceRows);
            if (_charSheet.Rows.Any())
            {
                TestProgram.IgxlWorkBk.AddCharSheet(FolderStructure.DirDevChar, _charSheet);
            }
            #endregion
            AddControlBinFlagsToMainInitEnableWd();
        }

        internal override List<PatSet> GenPatSets(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            var patSets = new List<PatSet>();
            foreach (BinCutFinalInstanceRow binCutFinalInstanceRow in binCutFinalInstanceRows)
            {
                if (!string.IsNullOrEmpty(binCutFinalInstanceRow.BinCutInstanceRow.Char))
                {
                    if (binCutFinalInstanceRow.InitPatSet != null)
                    {
                        patSets.Add(binCutFinalInstanceRow.InitPatSet);
                    }
                    if (binCutFinalInstanceRow.PatSet != null)
                    {
                        patSets.Add(binCutFinalInstanceRow.PatSet);
                    }
                    continue;
                }
                if (binCutFinalInstanceRow.PatternList.Count < 2 || !string.IsNullOrEmpty(binCutFinalInstanceRow.BinCutInstanceRow.Char))
                {
                    continue;
                }

                var patSet = new PatSet { PatSetName = binCutFinalInstanceRow.PatSetName };

                foreach (string pattern in binCutFinalInstanceRow.PatternList)
                {
                    var patSetRow = new PatSetRow { File = pattern, Burst = "No", Comment = "SheetName: " + binCutFinalInstanceRow.BinCutInstanceRow.SheetName + ", Row: " + binCutFinalInstanceRow.BinCutInstanceRow.RowNum };
                    patSet.AddRow(patSetRow);
                }
                if (!patSets.Exists(x => x.PatSetName.Equals(patSet.PatSetName)))
                {
                    patSets.Add(patSet);
                }
            }
            return patSets;
        }

        private SubFlowSheet GenCpmMainFlow(List<SubFlowSheet> subflows, ref InstanceRows instance)
        {
            string sheetName = "Flow_CPM";
            var mainFlow = new SubFlowSheet("Flow_CPM", TestPlanStatic.CpmInstanceSheets.First().SheetName);
            mainFlow.AddHeaderRow(sheetName);
            mainFlow.AddRows(AddInitFlags());

            foreach (SubFlowSheet subflow in subflows)
            {
                var flowRow = new FlowRow { Opcode = OpCode.Call, Parameter = subflow.Name };
                mainFlow.AddRow(flowRow);
            }
            mainFlow.AddFooterRow(sheetName);
            mainFlow.AddReturnRow();
            instance.AddHeader(sheetName);
            instance.AddFooter(sheetName);

            return mainFlow;
        }

        private List<FlowRow> AddInitFlags()
        {
            var flowRows = new List<FlowRow>();
            foreach (string flag in _allFlags.Distinct())
            {
                if (string.IsNullOrEmpty(flag) || flag.Equals("F_CPM_Default_Trim", StringComparison.OrdinalIgnoreCase) || flag.Equals("F_CPM_Alt_Trim", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                flowRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = flag });
            }
            flowRows.Add(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = "F_CPM_Default_Trim", Job = "CP1" });
            flowRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = "F_CPM_Alt_Trim", Job = "CP1" });
            return flowRows;
        }

        public override List<FlowRow> GenFlowBinTable(List<BinCutFinalInstanceRow> group, SubFlowSheet subFlowSheet = null, List<HarvestingTruthTableSheet> truthTables = null)
        {
            var flowRows = new List<FlowRow>();
            List<FlowRow> binFlowRows = GetBinFlowRows(group);
            foreach (FlowRow binFlowRow in binFlowRows)
            {
                (FlowRow ifRow, FlowRow endIfRow) = GetIfControlBinRow(binFlowRow);
                flowRows.Add(ifRow);
                flowRows.Add(binFlowRow);
                flowRows.Add(endIfRow);
            }
            return flowRows;
        }

        internal override List<AtpgBinOutData> GetBinOutDatas(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            List<AtpgBinOutData> resultsList = new List<AtpgBinOutData>();
            foreach (BinCutFinalInstanceRow binCutFinalInstanceRow in binCutFinalInstanceRows)
            {
                if (!binCutFinalInstanceRow.PatternList.Any())
                {
                    continue;
                }

                string pattern = binCutFinalInstanceRow.PayloadList == null || !binCutFinalInstanceRow.PayloadList.Any() ? binCutFinalInstanceRow.PatternList.First() : binCutFinalInstanceRow.PayloadList.First();
                string payloadType = GetPayloadType(pattern);
                List<string> failActionFlags = binCutFinalInstanceRow.GetInstanceFailFlags().Any() ?
                    binCutFinalInstanceRow.GetInstanceFailFlags() :
                    binCutFinalInstanceRow.GetDefaultFailFlags(payloadType, binCutFinalInstanceRow.PerformanceMode);

                string binOutStage = binCutFinalInstanceRow.GetBinOutStage();
                bool isByPassBinOut = binCutFinalInstanceRow.BinCutInstanceRow.IsBypassBinOut;

                resultsList.AddRange(failActionFlags.Select(x => new AtpgBinOutData("", "", "OR", binOutStage, isByPassBinOut, false, false, [x])));
            }
            return resultsList;
        }

        protected override List<FlowRow> GenFlowBody(IGrouping<string, BinCutFinalInstanceRow> group)
        {
            var flowRows = new List<FlowRow>();
            foreach (BinCutFinalInstanceRow dataRow in group)
            {
                Dictionary<string, string> specialSettings = dataRow.BinCutInstanceRow.GetSpecialSettings();
                if (specialSettings != null)
                {
                    if (specialSettings.ContainsKey("OverlayName"))
                    {
                        flowRows.AddRange(CreateOverlay(dataRow));
                    }
                    else if (specialSettings.ContainsKey("Pass"))
                    {
                        flowRows.AddRange(FlowWithDecision(dataRow));
                    }
                    else
                    {
                        flowRows.AddRange(WriteNoPatternRow(dataRow));
                    }
                    continue;
                }
                if (!string.IsNullOrEmpty(dataRow.BinCutInstanceRow.Char))
                {
                    flowRows.Add(CreateChar(dataRow));
                }
                List<FlowRow> testFlowRows = new List<FlowRow>();
                testFlowRows.Add(GetTestRow(dataRow, dataRow.PatSetName));
                if (!dataRow.BinCutInstanceRow.IsBypassBinOut)
                {
                    testFlowRows.Add(GetControlBinRow(dataRow));
                }
                List<FlowRow> ifFlowRows = new ScanNonBinCutInstance().GetIfFlowRows(dataRow, testFlowRows, null);
                flowRows.AddRange(ifFlowRows);
            }
            return flowRows;
        }

        private FlowRows FlowWithDecision(BinCutFinalInstanceRow row)
        {
            var flowRows = new FlowRows { GetTestRow(row, row.PatSetName) };
            Dictionary<string, string> specialSettings = row.BinCutInstanceRow.GetSpecialSettings();
            string devCon = row.BinCutInstanceRow.FailFlag;
            string job = row.BinCutInstanceRow.JobTestStage;
            foreach (KeyValuePair<string, string> setting in specialSettings)
            {
                string flag = setting.Value;
                if (!setting.Key.Equals("Pass", StringComparison.OrdinalIgnoreCase) && !setting.Key.Equals("Fail", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string condition = setting.Key.Equals("Pass", StringComparison.OrdinalIgnoreCase) ? OpCode.FlagFalse : OpCode.FlagTrue;
                flowRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = flag, Job = job });
                flowRows.Add(new FlowRow { Opcode = OpCode.FlagTrue, DeviceCondition = condition, Parameter = flag, DeviceName = devCon, Job = job });
            }
            return flowRows;
        }

        private FlowRow CreateChar(BinCutFinalInstanceRow row)
        {
            var flowRow = new FlowRow
            {
                Job = row.GetJob(),
                Enable = row.GetEnable(),
                Opcode = row.NopByEnableWord ? OpCode.Nop : OpCode.Characterize,
                Parameter = row.PatSetName + " " + row.BinCutInstanceRow.Char.Split('=').First().Trim()
            };
            return flowRow;
        }

        private List<FlowRow> CreateOverlay(BinCutFinalInstanceRow row)
        {
            var flowRows = new List<FlowRow>
            {
                new FlowRow{Opcode = row.NopByEnableWord ? OpCode.Nop : OpCode.Test,Parameter = row.BinCutInstanceRow.PatSetNameOrange ,Job = row.BinCutInstanceRow.JobTestStage},
                new FlowRow{Opcode = row.NopByEnableWord ? OpCode.Nop : OpCode.UseLimit, Parameter = row.BinCutInstanceRow.PatSetNameOrange + " CPMS_Margin_Check_LVCC", Job = row.BinCutInstanceRow.JobTestStage},
                new FlowRow { Opcode = row.NopByEnableWord ? OpCode.Nop : OpCode.UseLimit, Parameter = row.BinCutInstanceRow.PatSetNameOrange + " CPMS_Margin_Check_Func" ,Job = row.BinCutInstanceRow.JobTestStage}
            };

            return flowRows;
        }

        private List<FlowRow> WriteNoPatternRow(BinCutFinalInstanceRow row)
        {
            var flowRows = new List<FlowRow>
            {
                new FlowRow
                {
                    Opcode = row.NopByEnableWord ? OpCode.Nop : OpCode.Test,
                    Parameter = row.BinCutInstanceRow.PatSetNameOrange,
                    Job = row.GetJob(),
                    Enable = row.GetEnable(),
                    FailAction = row.BinCutInstanceRow.FailFlag
                }
            };
            return flowRows;
        }

        internal override FlowRow GetTestRow(BinCutFinalInstanceRow row, string parameter = "")
        {
            FlowRow flowRow = base.GetTestRow(row, parameter);
            List<string> failFlags = new List<string>();
            failFlags.AddRange(flowRow.FailAction.Split(','));
            failFlags.AddRange(flowRow.PassAction.Split(','));
            failFlags.Remove($"F_{parameter}");
            _allFlags.AddRange(failFlags);
            return flowRow;
        }

        internal override void SetPatSetSheet(List<PatSet> patSets)
        {
            if (patSets.Count > 0)
            {
                var patSetSheet = new PatSetSheet("PatSets_CPM");
                patSetSheet.AddRows(patSets);
                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirCpm, patSetSheet);
            }
        }

        private BinTableSheet GenCpmBinTable(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            var binTableSheet = new BinTableSheet(BinTableCpm);
            binTableSheet.AddRows(GetBinTableRows(binCutFinalInstanceRows));
            return binTableSheet;
        }
    }
}
