using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.PostAction.GenGroupFlowSheet;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Business;
using Automation.GenerateIgxl.PostAction.GenMainFlow.PreCheck;
using Automation.GenerateIgxl.PostAction.Relay;
using Automation.GenerateIgxl.PostAction.TempMon.Enums;
using Automation.PreCheck.AllParaData;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Basic;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using OfficeOpenXml;

using TestPlanLib;
using TestPlanLib.Basic;
using TestPlanLib.BinNumber;
using TestPlanLib.Concurrent;
using TestPlanLib.Singleton;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow
{
    public class MainFlowMain
    {
        public static readonly Regex RegexFlow = new Regex("^Flow_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        public const string SubProgramMainFlow = "Main_Flow_Sub";
        public const string T0txRoomProgramMainFlow = "Main_Flow_T0TX_Room";
        public const string T0txHotProgramManFlow = "Main_Flow_T0TX_Hot";
        public const string FlowMainStart = "Flow_Main_Start";
        public const string FlowMainEnd = "Flow_Main_End";

        public static List<string> SubProgramFlowList = new List<string>();

        private const string ConIds = "IDS";
        private const string ConSpi = "Spi";

        private const string ConSetDevice = "set-device";
        private const string ConSpiromMain = "Flow_Table_Write_SPIROM_Main";
        private const string ConSpirom32M = "Flow_Table_Write_SPIROM_32M";
        private const string ConFlowTableMainInitEnableWd = "Flow_Table_Main_Init_EnableWd";
        private const string ConFlowTableMainInitFlows = "Flow_Table_Main_Init_Flows";
        private const string ConFlowTablenWire = "Flow_nWire_";
        private const string ConVddBiningFlow = "Flow_Vddbinning";
        private const string SuffixShmooFlowSheet = "_CZ2";
        private const string ConKeywordTd = "Td";

        private const string ConCommentUnable = "Unable generated from Current Input files";
        private const string ConCommentUnableFromUi = "Disabled by User Interface";
        private const string ConCommentMissing = "Missing this Flow during AutoGen";
        private const string ConCommentUnableByTool = "Disabled by T-AutoGen";
        private const string ConCommentDidnTSpecifyInFlowMainSheetTestplan = "Didn't specify in Flow_Main sheet(TestPlan)!";
        private const string InitializeConst = "Initialize";
        private readonly List<SheetNameMappingTable> _sheetNameMappingTables = new List<SheetNameMappingTable>();
        private readonly VariableInitTable _variableInitTable;
        private readonly Dictionary<string, SubFlowSheet> _subFlowSheets;
        private readonly Dictionary<string, InstanceSheet> _insSheets;
        private readonly InstanceSheet _commonInstanceSheet;
        private readonly List<RelayItemNew> _relays;
        private readonly List<SubFlowSheet> _groupSubFlow = new List<SubFlowSheet>();
        private readonly List<string> _t0txExcludeFlow = new List<string> { "Flow_Efuse_Post", "Flow_EVS" };

        public MainFlowMain(VariableInitTable variableInitTable, Dictionary<string, SubFlowSheet> subFlowSheets, Dictionary<string, InstanceSheet> insSheets, List<RelayItemNew> relays = null)
        {
            _variableInitTable = variableInitTable;
            _subFlowSheets = subFlowSheets;
            _insSheets = insSheets;
            _relays = relays;
            _commonInstanceSheet = _insSheets.FirstOrDefault(x => x.Value.Name.Equals("TestInst_Common")).Value;
        }

        public (List<string>, List<MainFlow>, List<SubFlowSheet>) GenMainFlow(MainFlowSheet mainFlowSheet)
        {
            (List<string> allFlags, List<MainFlow> mainFlows, List<SubFlowSheet> allFlows) = GenMainFlowsNew(ref mainFlowSheet);
            return (allFlags, mainFlows, allFlows);
        }

        private void DeferredBinoutProcess(MainFlowSheet mainFlowSheet)
        {
            if (LocalSpecs.IsModuleIncluded("Defer"))
            {
                var deferredBinOutProcessor = new DeferredBinOutProcessor();
                Dictionary<string, HashSet<string>> deferredBinoutJobs =
                    deferredBinOutProcessor.GetDeferredBinoutJobs(mainFlowSheet);
                Dictionary<string, List<SubFlowSheet>> deferredBinoutMapFlows =
                    deferredBinOutProcessor.GetDeferredBinoutMaps(mainFlowSheet, deferredBinoutJobs)
                    .ToDictionary(x => x.Key, y => y.Value.SelectMany(z => GetSubFlowsBySequence(z)).ToList());
                Dictionary<string, SubFlowSheet> deferredSubFlows =
                    deferredBinOutProcessor.GetDeferredBinoutFlows(_commonInstanceSheet, deferredBinoutMapFlows, deferredBinoutJobs, mainFlowSheet.Jobs);
                foreach (string deferredFlowName in deferredSubFlows.Keys)
                {
                    _sheetNameMappingTables.Add(new SheetNameMappingTable { SheetName = [deferredFlowName], SourceSheet = deferredFlowName, SubFlowSheet = deferredFlowName });
                }
            }
        }

        public (List<string>, List<MainFlow>, List<SubFlowSheet>) GenMainFlowsNew(ref MainFlowSheet mainFlowSheet)
        {
            List<MainFlow> mainFlows = new List<MainFlow>();
            List<string> allFlags = new List<string>();
            List<string> allJobs = mainFlowSheet.Jobs;
            List<SubFlowSheet> subprogramFlows = new List<SubFlowSheet>();
            List<MainFlow> mainflowSubs = new List<MainFlow>();
            SetupSheetNameMappingTable();
            CreateChipLetSubFlow(mainFlowSheet.Rows);
            DeferredBinoutProcess(mainFlowSheet);

            foreach (MainFlowBase mainFlow in mainFlowSheet.Rows)
            {
                if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any() && !TestPlanStatic.TestProgramDefSheet.JobMappingDic.ContainsKey(mainFlow.JobName.ToUpper()))
                {
                    continue;
                }
                if (!mainFlow.MainFlowName.Equals("Main_Flow_Sub"))
                {
                    MainFlow flow = GenMainFlow(mainFlow);
                    if (mainFlow.MainFlowName.ContainsIgnoreCase("T0TX"))
                    {
                        if (flow.Rows.Exists(x => x.Opcode.Equals(OpCode.Call) && _t0txExcludeFlow.Exists(y => y.Equals(x.Parameter, StringComparison.CurrentCultureIgnoreCase))))
                        {
                            List<FlowRow> targetRow = flow.Rows.FindAll(x => x.Opcode.Equals(OpCode.Call) && _t0txExcludeFlow.Exists(y => y.Equals(x.Parameter, StringComparison.CurrentCultureIgnoreCase)));
                            foreach (FlowRow row in targetRow)
                            {
                                row.Opcode = OpCode.Nop;
                            }
                        }
                    }
                    mainFlows.Add(flow);
                }
                if (mainFlowSheet.SubFlowCol != -1 && mainFlow.MainFlowName.Equals("Main_Flow_Sub"))
                {
                    (SubFlowSheet mainFlowSubBase, List<SubFlowSheet> flows) = GenerateMainFlowBySuBProgram(mainFlow, allJobs);
                    subprogramFlows.AddRange(flows);
                    foreach (SubprogramSetting subprogramSetting in TestPlanStatic.SubprogramMappingSheet.SubprogramSettings)
                    {
                        string mainflowSubName = "Main_Flow_" + subprogramSetting.SubprogramName;
                        MainFlow mainflowSub = new MainFlow(mainflowSubName);
                        foreach (FlowRow row in mainFlowSubBase.Rows)
                        {
                            FlowRow rowCopy = row.Copy();
                            if (row.Opcode.Equals(OpCode.Call) && !subprogramSetting.EnableSubFlows.Contains(row.Parameter))
                            {
                                rowCopy.Opcode = OpCode.Nop;
                            }
                            mainflowSub.Rows.Add(rowCopy);
                        }
                        mainflowSubs.Add(mainflowSub);
                    }
                    mainFlows.AddRange(mainflowSubs);
                }
                AddErrorReport(mainFlow);
                foreach (string flag in mainFlow.SequencesNew.FindAll(x => x.Enable).Select(x => x.FailFlag))
                {
                    if (!string.IsNullOrEmpty(flag) && !allFlags.Exists(x => x.Equals(flag)))
                    {
                        allFlags.Add(flag);
                    }
                }
            }

            return (allFlags, mainFlows, subprogramFlows);
        }

        private bool PreCheck()
        {
            PreCheckMainFlow preCheckMainFlow = new PreCheckMainFlow(EpWorkbook.TestPlanWorkbook);
            preCheckMainFlow.SetSheetName(NeededSheets.MainFlow);
            bool lBResult = preCheckMainFlow.CheckMain();
            return lBResult;
        }

        public void WorkFlow(PostActionParaData paraData)
        {
            ExcelWorksheet genMainFlow = EpWorkbook.TestPlanWorkbook.Worksheets["GenMainFlow"];

            if (genMainFlow != null)
            {
                ReadFlowSheet flowsheets = new ReadFlowSheet();
                SubFlowSheet mainFlow = flowsheets.ReadSheet(genMainFlow);
                foreach (List<string> jobList in LocalSpecs.JobMap.Values)
                {
                    foreach (string jobName in jobList)
                    {
                        MainFlow flow = new MainFlow(NeededSheets.MainFlow + "_" + jobName) { Rows = mainFlow.Rows };
                        TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);
                    }
                }
            }
            else if (PreCheck())
            {
                genMainFlow = EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.MainFlow];
                MainFlowGen(genMainFlow);
            }
        }

        private void GenOtherTestItemToMainFlow(ref MainFlow flowSheet)
        {
            string job = flowSheet.Name;
            MainFlow tmpSubFlowSheet = new MainFlow(flowSheet.Name);
            foreach (FlowRow flow in flowSheet.Rows)
            {
                if (flow?.Parameter != null)
                {
                    #region SelSram

                    if (flow.Parameter.Equals("Flow_SelSram", StringComparison.OrdinalIgnoreCase))
                    {
                        flow.Enable = "SelSram";
                        tmpSubFlowSheet.AddRow(flow);
                        continue;
                    }

                    #endregion

                    #region Conti add EVS blank check on CP1/CP2

                    if (flow.Parameter.Equals("Flow_DC_Conti", StringComparison.OrdinalIgnoreCase))
                    {

                        tmpSubFlowSheet.AddRow(flow);
                        if (Regex.IsMatch(job, "CP1|CP2", RegexOptions.IgnoreCase))
                        {
                            FlowRow evsBc =
                                flowSheet.Rows.FirstOrDefault(
                                    p =>
                                        p.Parameter.Equals("Flow_eFuse_BC_4_EVS", StringComparison.OrdinalIgnoreCase));
                            if (evsBc != null)
                            {
                                tmpSubFlowSheet.AddRow(new FlowRow { Opcode = "call", Parameter = "Flow_eFuse_BC_4_EVS" });
                            }
                        }
                        continue;
                    }
                    if (flow.Parameter.Equals("Flow_eFuse_BC_4_EVS", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    #endregion

                    #region Fuse EVS

                    string flowName = Regex.Match(flow.Parameter, @"flow_(?<name>\w+)", RegexOptions.IgnoreCase).Groups["name"].Value;
                    if (flowName.Equals("EVS", StringComparison.OrdinalIgnoreCase) &&
                        flowName.IndexOf("efuse", StringComparison.OrdinalIgnoreCase) == -1)
                    {
                        //If F_ECID_All_BC is False?
                        //If F_CFG_All_BC is False?
                        tmpSubFlowSheet.AddRow(new FlowRow
                        {
                            Opcode = "if",
                            Parameter =
                               Regex.IsMatch(job, "CP1", RegexOptions.IgnoreCase)
                                    ? "!F_ECID_All_BC"
                                    : "!F_CFG_All_BC"
                        });
                        tmpSubFlowSheet.AddRow(flow);
                        tmpSubFlowSheet.AddRow(new FlowRow
                        {
                            Opcode = "endif",
                            Parameter = ""
                        });
                        TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(new List<string> { "F_ECID_All_BC", "F_CFG_All_BC" }, "Efuse", FolderStructure.DirMain);
                    }
                    else
                    {
                        tmpSubFlowSheet.AddRow(flow);
                    }

                    #endregion

                }

            }
            tmpSubFlowSheet.JobNames = flowSheet.JobNames;
            tmpSubFlowSheet.HasConcurrent = flowSheet.HasConcurrent;
            flowSheet = tmpSubFlowSheet;
        }

        private void AppendRelayToMainFlow(ref MainFlow flowSheet)
        {
            MainFlow tmpSubFlowSheet = new MainFlow(flowSheet.Name);
            string relayModule = "";
            foreach (FlowRow flow in flowSheet.Rows)
            {
                if (flow?.Parameter != null)
                {
                    if (!string.IsNullOrEmpty(relayModule))
                    {
                        if (Regex.IsMatch(flow.Parameter, relayModule, RegexOptions.IgnoreCase))
                        {//if match to current module, do nothing 
                        }
                        else
                        {
                            relayModule = "";
                        }
                    }
                    else if (flow.Opcode.Equals("call", StringComparison.OrdinalIgnoreCase))
                    {
                        RelayItemNew relayOn = _relays.FirstOrDefault(p => Regex.IsMatch(flow.Parameter, p.Module, RegexOptions.IgnoreCase));
                        if (relayOn != null)
                        {
                            relayModule = relayOn.Module;
                            tmpSubFlowSheet.AddRow(new FlowRow
                            {
                                Opcode = "Test",
                                Parameter = $"Relay_{relayOn.Module}"
                            });
                        }
                    }
                    tmpSubFlowSheet.AddRow(flow);

                }

            }
            tmpSubFlowSheet.HasConcurrent = flowSheet.HasConcurrent;
            tmpSubFlowSheet.JobNames = flowSheet.JobNames;
            flowSheet = tmpSubFlowSheet;

        }

        internal void MainFlowGen(ExcelWorksheet mainFlowWorksheet)
        {
            MainFlowSheetReader mainFlowSheetReader = new MainFlowSheetReader();
            MainFlowSheet mainFlowSheet = mainFlowSheetReader.ReadSheet(mainFlowWorksheet);
            if (mainFlowSheet.Rows.Any())
            {
                List<MainFlow> flows = GenMainFlows(ref mainFlowSheet);
                foreach (MainFlow flow in flows)
                {
                    TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, flow);
                }
            }
            else
            {
                mainFlowSheet = TestPlanStatic.MainFlowSheet ?? TestPlanStatic.MainFlowSheet;
                (List<string> allFlags, List<MainFlow> mainflows, List<SubFlowSheet> flows) = GenMainFlow(mainFlowSheet);
                SubProgramFlowList = flows.Select(x => x.Name).ToList();
                AddGroupSubflowIntoProg();

                foreach (MainFlow mainflow in mainflows)
                {
                    if (mainflow.Name.StartsWith(SubProgramMainFlow))
                    {
                        TestProgram.SubProgIgxlWorkBk.AddMainFlowSheet(FolderStructure.DirSubProgram, mainflow);
                    }
                    else if (mainflow.Name.ContainsIgnoreCase("T0TX"))
                    {
                        TestProgram.T0TxIgxlWorkBk.AddMainFlowSheet(FolderStructure.DirT0Tx, mainflow);
                    }
                    else
                    {
                        TestProgram.IgxlWorkBk.AddMainFlowSheet(FolderStructure.DirMain, mainflow);
                    }
                }

                foreach (SubFlowSheet flow in flows)
                {
                    TestProgram.SubProgIgxlWorkBk.AddSubFlowSheet(FolderStructure.DirSubProgram, flow);
                }

                TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(allFlags, "MainFlow", FolderStructure.DirMain);
            }
        }

        public List<MainFlow> GenMainFlows(ref MainFlowSheet mainFlowSheet)
        {
            GenGroupSubFlowSheets(mainFlowSheet.Rows);

            List<MainFlowBase> mainFlows = new GroupFlowSheetMain().WorkFlow(_subFlowSheets, mainFlowSheet.Rows);
            List<MainFlow> flows = new List<MainFlow>();
            foreach (MainFlowBase mainFlow in mainFlows)
            {
                MainFlow flow = GenMainFlowSheet(mainFlow);
                GenOtherTestItemToMainFlow(ref flow);
                if (_relays != null && _relays.Count > 0)
                {
                    AppendRelayToMainFlow(ref flow);
                }
                flows.Add(flow);
            }

            return flows;
        }

        private void AddErrorReport(MainFlowBase flowMain)
        {
            foreach (FlowSequenceNew row in flowMain.SequencesNew.Where(x => x.Enable && !x.Found))
            {
                string subFlowErrorMsg = !string.IsNullOrEmpty(row.SubFlowName) ? $" (Sub Flow: \"{row.SubFlowName}\")" : "";
                ErrorReportManager.AddError(FlowMainErrorType.W_MismatchFlow_01, flowMain.SheetName, row.RowNum, flowMain.JobColIndex,
                    [row.SheetName, subFlowErrorMsg, row.Job]);
            }
        }

        private void AddGroupSubflowIntoProg()
        {
            foreach (SubFlowSheet subflow in _groupSubFlow)
            {
                SubFlowSheet reOrderRowByNum = new SubFlowSheet(subflow.Name);
                reOrderRowByNum.AddRows(subflow.Rows.OrderBy(x => x.RowNum).ToList());
                reOrderRowByNum.AddReturnRow();
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirCommon, reOrderRowByNum);
            }
        }

        private void CreateChipLetSubFlow(List<MainFlowBase> mainFlows)
        {
            foreach (MainFlowBase mainFlow in mainFlows)
            {
                foreach (string group in mainFlow.GetAllGroups.Distinct())
                {
                    if (!_groupSubFlow.Exists(x => x.Name.Equals(group, StringComparison.OrdinalIgnoreCase)))
                    {
                        _groupSubFlow.Add(new SubFlowSheet(group));
                        _sheetNameMappingTables.Add(new SheetNameMappingTable { SheetName = new List<string> { group }, SourceSheet = "T-Autogen", SubFlowSheet = group });
                    }
                }
            }
        }

        internal string GetHarvDecisionFlow(string type, string job)
        {
            SheetNameMappingTable sheets = _sheetNameMappingTables.Find(x => x.SourceSheet.StartsWith("HarvestingTruthTable", StringComparison.OrdinalIgnoreCase));
            if (sheets == null)
            {
                return "";
            }

            return sheets.SheetName.Find(x => x.Contains($"_{job}_{type}"));
        }

        internal void SetupSheetNameMappingTable()
        {
            foreach (SubFlowSheet subflow in _subFlowSheets.Values)
            {
                string sheetName = subflow.Name;
                string[] sourceSheetItems = subflow.SourceInfo.Name.Split(':');
                string sourceSheet = sourceSheetItems[0];
                string subFlowSheet = sourceSheetItems.Length > 1 ? sourceSheetItems[1] : "";
                SheetNameMappingTable targetSource = _sheetNameMappingTables.Find(x => x.SourceSheet.Equals(sourceSheet, StringComparison.OrdinalIgnoreCase) && x.SubFlowSheet.Equals(subFlowSheet, StringComparison.OrdinalIgnoreCase));
                if (targetSource == null)
                {
                    _sheetNameMappingTables.Add(new SheetNameMappingTable { SheetName = new List<string> { sheetName }, SourceSheet = sourceSheet, SubFlowSheet = subFlowSheet });
                }
                else
                {
                    targetSource.SheetName.Add(sheetName);
                }
            }
        }

        private (SubFlowSheet, List<SubFlowSheet>) GenerateMainFlowBySuBProgram(MainFlowBase flowMain, List<string> allJobs)
        {
            List<SubFlowSheet> flows = new List<SubFlowSheet>();

            MainFlow subProgramMainFlow = new MainFlow(SubProgramMainFlow);
            List<string> avalibleParts = flowMain.AvalibleParts;
            string currentModule = "";
            subProgramMainFlow.AddRow(GenAlarmRow());
            subProgramMainFlow.AddRow(GenErrorRow());
            FlowRow initTestTimeUtilityRow = GetInitTestTimeUtilityRow();
            if (initTestTimeUtilityRow != null)
            {
                subProgramMainFlow.AddRow(initTestTimeUtilityRow);
            }

            AddVariable(ref subProgramMainFlow, flowMain.JobName, allJobs);

            IEnumerable<IGrouping<string, FlowSequenceNew>> groups = flowMain.SequencesNew.ChunkBy(x => x.SubProgramFlowName);

            foreach (IGrouping<string, FlowSequenceNew> group in groups)
            {
                string name = group.Key;
                MainFlow flow = new MainFlow(name);
                foreach (FlowSequenceNew item in group)
                {
                    int startIdx = flow.Rows.Count;
                    string jobs = item.GetJobs();
                    FlowRow flowRow = DispatchItemToSubProgramFlow(item, flowMain, avalibleParts, jobs, ref flow);
                    SetTempMonFromFlowSequence(item, flowRow);

                    AddTestNameWidthRow(startIdx, currentModule, item.Module, ref flow);
                    currentModule = item.Module;
                }
                flow.AddRow(new FlowRow { Opcode = OpCode.Return });
                subProgramMainFlow.AddRow(new FlowRow { Opcode = OpCode.Call, Parameter = flow.Name });
                flows.Add(flow);
            }

            subProgramMainFlow.AddRow(new FlowRow { Opcode = OpCode.LimitsAll });
            subProgramMainFlow.AddRow(GenSetDevice());

            return (subProgramMainFlow, flows);
        }

        private FlowRow DispatchItemToSubProgramFlow(FlowSequenceNew item, MainFlowBase flowMain, List<string> avalibleParts, string jobs, ref MainFlow flow)
        {
            FlowRow flowRow = null;
            if (item.Source.Equals("IGXL"))
            {
                if (EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Exists(x => x.Name.Equals(item.OriSheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    flowRow = new FlowRow { Opcode = OpCode.Call, Parameter = item.OriSheetName, Enable = item.EnableWord, FailAction = item.FailFlag };
                    item.Found = true;
                    AddRowIntoFlow(flowRow, ref flow, item, jobs);
                }
                List<string> otherSheets = item.Comment.Split(',').ToList();
                foreach (string otherSheet in otherSheets)
                {
                    string sheetName = Path.GetFileNameWithoutExtension(otherSheet);
                    ExcelWorksheet targetSheet = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Find(x => x.Name.Equals(item.OriSheetName, StringComparison.OrdinalIgnoreCase));
                    if (targetSheet == null)
                    {
                        continue;
                    }

                    targetSheet.ExportWorkBook2Txt(FolderStructure.DirCommon);
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, sheetName);
                }
            }
            else if (item.Module.ToUpper().Equals("BINTABLE"))
            {
                if (!string.IsNullOrEmpty(item.SheetName) && !string.IsNullOrEmpty(item.Option))
                {
                    string opcode = OpCode.BinTable;
                    string parameter = item.SheetName;
                    List<string> flagList = item.Option.Split(',').ToList();
                    CreateBinTable(flagList, parameter);
                    flowRow = new FlowRow { Opcode = opcode, Parameter = parameter, Enable = item.EnableWord };
                    AddRowIntoFlow(flowRow, ref flow, item, jobs);
                }
            }
            else if (item.Module.ToUpper().Equals("OPCODE"))
            {
                if (!string.IsNullOrEmpty(item.SheetName))
                {
                    string opcode = item.SheetName;
                    string parameter = item.SubFlowName;
                    flowRow = new FlowRow { Opcode = opcode, Parameter = parameter, Enable = item.EnableWord, FailAction = item.FailFlag };
                    AddRowIntoFlow(flowRow, ref flow, item, jobs);
                }
            }
            else if (item.Module.Equals("CHAR", StringComparison.OrdinalIgnoreCase) && LocalSpecs.IsModuleIncluded("CHAR"))
            {
                if (!string.IsNullOrEmpty(item.SheetName) && !string.IsNullOrEmpty(item.SubFlowName))
                {
                    bool flowIsExist = false;
                    foreach (string customPath in LocalSpecs.CustomPath)
                    {
                        if (Directory.Exists(customPath) && !flowIsExist)
                        {
                            IEnumerable<string?> extraSubflow = IgxlSheetReaderHelpers.GetSheetsByType(customPath, EnumSheetType.DTFlowtableSheet).Select(Path.GetFileNameWithoutExtension);
                            flowIsExist = extraSubflow.ToList().Exists(x => x.Equals(item.SubFlowName, StringComparison.OrdinalIgnoreCase));
                        }
                    }
                    flowRow = new FlowRow { Opcode = flowIsExist ? OpCode.Call : OpCode.Nop, Parameter = item.SubFlowName, Enable = item.EnableWord, ColumnA = item.SheetName, FailAction = item.FailFlag };
                    AddRowIntoFlow(flowRow, ref flow, item, jobs);
                }
            }
            else if (item.Source.ToUpper().Equals("CUSTOM") && LocalSpecs.IsModuleIncluded("IGXL"))
            {
                bool flowIsExist = false;
                foreach (string customPath in LocalSpecs.CustomPath)
                {
                    if (Directory.Exists(customPath) && !flowIsExist)
                    {
                        IEnumerable<string?> extraSubflow = IgxlSheetReaderHelpers.GetSheetsByType(customPath, EnumSheetType.DTFlowtableSheet).Select(Path.GetFileNameWithoutExtension);
                        flowIsExist = extraSubflow.ToList().Exists(x => x.Equals(item.SheetName));
                    }
                }
                if (flowIsExist)
                {
                    flowRow = new FlowRow { Opcode = OpCode.Call, Parameter = item.SheetName, Enable = item.EnableWord, FailAction = item.FailFlag };
                    AddRowIntoFlow(flowRow, ref flow, item, jobs);
                }
            }
            else
            {
                flowRow = DispatchItemToSubProgramFlowBySheetName(item, flowMain, avalibleParts, jobs, ref flow);
            }
            return flowRow;
        }

        private FlowRow DispatchItemToSubProgramFlowBySheetName(FlowSequenceNew item, MainFlowBase flowMain, List<string> avalibleParts, string jobs, ref MainFlow flow)
        {
            FlowRow flowRow = null;
            if (item.SheetName.Equals(InitializeConst, StringComparison.CurrentCultureIgnoreCase))
            {
                ConvertToRealFlowName(item, ref flow, avalibleParts);
            }
            else if (item.SheetName.StartsWith("HarvestingTruthTable", StringComparison.OrdinalIgnoreCase))
            {
                string harvType = item.SheetName.Split('_').Last();
                foreach (string job in item.GetJobList())
                {
                    string harvSheetName = GetHarvDecisionFlow(harvType, flowMain.GetIgxlJobName(job));
                    if (string.IsNullOrEmpty(harvSheetName))
                    {
                        continue;
                    }

                    string part = "";
                    if (!_subFlowSheets.Values.ToList().Exists(x => x.Name.Equals(harvSheetName, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(item.Part) && avalibleParts != null && !string.Join(",", avalibleParts).Equals(item.Part, StringComparison.OrdinalIgnoreCase))
                    {
                        part = item.Part;
                    }

                    flowRow = new FlowRow
                    {
                        Opcode = OpCode.Call,
                        Parameter = harvSheetName,
                        Enable = item.EnableWord,
                        Part = part,
                        FailAction = item.FailFlag
                    };
                    AddRowIntoFlow(flowRow, ref flow, item, job);
                }
            }
            else if (item.SheetName.Equals("Relay", StringComparison.OrdinalIgnoreCase))
            {
                InstanceRow instanceRow = _commonInstanceSheet?.Rows.FirstOrDefault(x => x.TestName.Equals(item.SubFlowName, StringComparison.OrdinalIgnoreCase));
                if (instanceRow != null)
                {
                    flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = item.SubFlowName, Enable = item.EnableWord, FailAction = item.FailFlag };
                    AddRowIntoFlow(flowRow, ref flow, item, jobs);
                    foreach (FlowRow bintableRow in CreateBinTable(item.FailFlag.Split(',').ToList()))
                    {
                        AddRowIntoFlow(bintableRow, ref flow, item, jobs);
                    }
                }
            }
            else if (item.SheetName.Equals("UF_Instance", StringComparison.OrdinalIgnoreCase))
            {
                InstanceSheet instanceUf = _insSheets.FirstOrDefault(x => x.Value.Name.Equals("TestInst_UF")).Value;
                if (instanceUf == null)
                {
                    return null;
                }

                InstanceRow instanceRow = instanceUf.Rows.FirstOrDefault(x => x.TestName.Equals(item.SubFlowName));
                if (instanceRow != null)
                {
                    flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = item.SubFlowName, Enable = item.EnableWord, FailAction = item.FailFlag };
                    AddRowIntoFlow(flowRow, ref flow, item, jobs);
                    foreach (FlowRow bintableRow in CreateBinTable(item.FailFlag.Split(',').ToList(), "", item.BintableEnableWord))
                    {
                        AddRowIntoFlow(bintableRow, ref flow, item, jobs);
                    }
                }
            }
            else
            {
                ConvertToRealFlowName(item, ref flow, avalibleParts, jobs);
            }
            return flowRow;
        }

        private void SetTempMonFromFlowSequence(FlowSequenceNew sequenceItem, FlowRow flowRow)
        {
            if (flowRow == null)
            {
                return;
            }

            if (TempMonService.TryGetTempMonSyntax(sequenceItem.Option, out string mode, out EnumCondition condition))
            {
                if (flowRow.Opcode == OpCode.Call)
                {
                    List<string> subFlowSheets = TestProgram.IgxlWorkBk.GetSubFlows(flowRow, OpCode.Call).Select(x => x.Parameter).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    subFlowSheets.ForEach(flowSheetName => TempMonService.SetTempMon(LocalSpecs.TempMonDatas, mode, condition, EnumType.Flow, flowSheetName));
                }
                else if (flowRow.Opcode == OpCode.Test)
                {
                    TempMonService.SetTempMon(LocalSpecs.TempMonDatas, mode, condition, EnumType.Instance, flowRow.Parameter);
                }
            }
        }

        internal MainFlow GenMainFlow(MainFlowBase flowMain)
        {
            MainFlow mainFlow = new MainFlow(flowMain.MainFlowName);
            string jobName = flowMain.JobName;
            if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any() && TestPlanStatic.TestProgramDefSheet.JobMappingDic.ContainsKey(jobName))
            {
                jobName = TestPlanStatic.TestProgramDefSheet.JobMappingDic[jobName];
            }
            mainFlow.JobNames = new List<string> { jobName };
            List<string> avalibleParts = flowMain.AvalibleParts;
            string currentModule = "";
            mainFlow.AddRow(GenAlarmRow());
            mainFlow.AddRow(GenErrorRow());
            FlowRow initTestTimeUtilityRow = GetInitTestTimeUtilityRow();
            if (initTestTimeUtilityRow != null)
            {
                mainFlow.AddRow(initTestTimeUtilityRow);
            }

            foreach (FlowSequenceNew sequence in flowMain.SequencesNew)
            {
                if (!sequence.Enable)
                {
                    continue;
                }

                int startIdx = mainFlow.Rows.Count;
                FlowRow flowRow = DispatchSequenceToMainFlow(sequence, flowMain, avalibleParts, ref mainFlow);
                SetTempMonFromFlowSequence(sequence, flowRow);
                AddTestNameWidthRow(startIdx, currentModule, sequence.Module, ref mainFlow);
                currentModule = sequence.Module;
            }
            mainFlow.AddRow(new FlowRow { Opcode = OpCode.LimitsAll });
            mainFlow.AddRow(GenSetDevice());
            //mainFlow.AddReturnRow();
            return mainFlow;
        }

        private FlowRow DispatchSequenceToMainFlow(FlowSequenceNew sequence, MainFlowBase flowMain, List<string> avalibleParts, ref MainFlow mainFlow)
        {
            FlowRow flowRow = null;
            if (sequence.Source.Equals("IGXL"))
            {
                if (EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Exists(x => x.Name.Equals(sequence.OriSheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    flowRow = new FlowRow { Opcode = OpCode.Call, Parameter = sequence.OriSheetName, Enable = sequence.EnableWord, FailAction = sequence.FailFlag };
                    sequence.Found = true;
                    AddRowIntoFlow(flowRow, ref mainFlow, sequence);
                }
                List<string> otherSheets = sequence.Comment.Split(',').ToList();
                foreach (string otherSheet in otherSheets)
                {
                    string sheetName = Path.GetFileNameWithoutExtension(otherSheet);
                    ExcelWorksheet targetSheet = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Find(x => x.Name.Equals(sequence.OriSheetName, StringComparison.OrdinalIgnoreCase));
                    if (targetSheet == null)
                    {
                        continue;
                    }

                    targetSheet.ExportWorkBook2Txt(FolderStructure.DirCommon);
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, sheetName);
                }
            }
            else if (sequence.Module.ToUpper().Equals("BINTABLE"))
            {
                if (!string.IsNullOrEmpty(sequence.SheetName) && !string.IsNullOrEmpty(sequence.Option))
                {
                    string opcode = OpCode.BinTable;
                    string parameter = sequence.SheetName;
                    List<string> flagList = sequence.Option.Split(',').ToList();
                    CreateBinTable(flagList, parameter);
                    flowRow = new FlowRow { Opcode = opcode, Parameter = parameter, Enable = sequence.EnableWord };
                    AddRowIntoFlow(flowRow, ref mainFlow, sequence);
                }
            }
            else if (sequence.Module.ToUpper().Equals("OPCODE"))
            {
                if (!string.IsNullOrEmpty(sequence.SheetName))
                {
                    string opcode = sequence.SheetName;
                    string parameter = sequence.SubFlowName;
                    flowRow = new FlowRow { Opcode = opcode, Parameter = parameter, Enable = sequence.EnableWord, FailAction = sequence.FailFlag };
                    AddRowIntoFlow(flowRow, ref mainFlow, sequence);
                }
            }
            else if (sequence.Module.Equals("CHAR", StringComparison.OrdinalIgnoreCase) && LocalSpecs.IsModuleIncluded("CHAR"))
            {
                if (!string.IsNullOrEmpty(sequence.SheetName) && !string.IsNullOrEmpty(sequence.SubFlowName))
                {
                    bool flowIsExist = false;
                    foreach (string customPath in LocalSpecs.CustomPath)
                    {
                        if (Directory.Exists(customPath) && !flowIsExist)
                        {
                            IEnumerable<string?> extraSubflow = IgxlSheetReaderHelpers.GetSheetsByType(customPath, EnumSheetType.DTFlowtableSheet).Select(Path.GetFileNameWithoutExtension);
                            flowIsExist = extraSubflow.ToList().Exists(x => x.Equals(sequence.SubFlowName, StringComparison.OrdinalIgnoreCase));
                        }
                    }
                    flowRow = new FlowRow { Opcode = flowIsExist ? OpCode.Call : OpCode.Nop, Parameter = sequence.SubFlowName, Enable = sequence.EnableWord, ColumnA = sequence.SheetName, FailAction = sequence.FailFlag };
                    AddRowIntoFlow(flowRow, ref mainFlow, sequence);
                }
            }
            else if (sequence.Source.ToUpper().Equals("CUSTOM") && LocalSpecs.IsModuleIncluded("IGXL"))
            {
                bool flowIsExist = false;
                foreach (string customPath in LocalSpecs.CustomPath)
                {
                    if (Directory.Exists(customPath) && !flowIsExist)
                    {
                        IEnumerable<string?> extraSubflow = IgxlSheetReaderHelpers.GetSheetsByType(customPath, EnumSheetType.DTFlowtableSheet).Select(Path.GetFileNameWithoutExtension);
                        flowIsExist = extraSubflow.ToList().Exists(x => x.Equals(sequence.SheetName));
                    }
                }
                if (flowIsExist)
                {
                    flowRow = new FlowRow { Opcode = OpCode.Call, Parameter = sequence.SheetName, Enable = sequence.EnableWord, FailAction = sequence.FailFlag };
                    AddRowIntoFlow(flowRow, ref mainFlow, sequence);
                }
            }
            else
            {
                flowRow = DispatchSequenceToMainFlowBySheetName(sequence, flowMain, avalibleParts, ref mainFlow);
            }
            return flowRow;
        }

        private FlowRow DispatchSequenceToMainFlowBySheetName(FlowSequenceNew sequence, MainFlowBase flowMain, List<string> avalibleParts, ref MainFlow mainFlow)
        {
            FlowRow flowRow = null;
            if (sequence.SheetName.Equals(InitializeConst))
            {
                if (mainFlow.Name.Equals("Main_Flow_T0TX_Room", StringComparison.CurrentCultureIgnoreCase))
                {
                    AddVariable(ref mainFlow, "CP1");
                    AddVariable(ref mainFlow, "FT1");
                }
                else if (mainFlow.Name.Equals("Main_Flow_T0TX_Hot", StringComparison.CurrentCultureIgnoreCase))
                {
                    AddVariable(ref mainFlow, "CP2");
                    AddVariable(ref mainFlow, "FT2");
                }
                else
                {
                    AddVariable(ref mainFlow, flowMain.JobName);
                }
                ConvertToRealFlowName(sequence, ref mainFlow, avalibleParts);
            }
            else if (sequence.SheetName.StartsWith("HarvestingTruthTable", StringComparison.OrdinalIgnoreCase))
            {
                string harvType = sequence.SheetName.Split('_').Last();
                string harvSheetName = GetHarvDecisionFlow(harvType, flowMain.IgxlJobName);
                if (string.IsNullOrEmpty(harvSheetName))
                {
                    return null;
                }

                string part = "";
                if (!_subFlowSheets.Values.ToList().Exists(x => x.Name.Equals(harvSheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    return null;
                }

                if (!string.IsNullOrEmpty(sequence.Part) && avalibleParts != null && !string.Join(",", avalibleParts).Equals(sequence.Part, StringComparison.OrdinalIgnoreCase))
                {
                    part = sequence.Part;
                }

                flowRow = new FlowRow
                {
                    Opcode = OpCode.Call,
                    Parameter = harvSheetName,
                    Enable = sequence.EnableWord,
                    Part = part,
                    FailAction = sequence.FailFlag
                };
                AddRowIntoFlow(flowRow, ref mainFlow, sequence);
            }
            else if (sequence.SheetName.Equals("Relay", StringComparison.OrdinalIgnoreCase))
            {
                InstanceRow instanceRow = _commonInstanceSheet?.Rows.FirstOrDefault(x => x.TestName.Equals(sequence.SubFlowName, StringComparison.OrdinalIgnoreCase));
                if (instanceRow != null)
                {
                    flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = sequence.SubFlowName, Enable = sequence.EnableWord, FailAction = sequence.FailFlag };
                    AddRowIntoFlow(flowRow, ref mainFlow, sequence);
                    foreach (FlowRow bintableRow in CreateBinTable(sequence.FailFlag.Split(',').ToList()))
                    {
                        AddRowIntoFlow(bintableRow, ref mainFlow, sequence);
                    }
                }
            }
            else if (sequence.SheetName.Equals("UF_Instance", StringComparison.OrdinalIgnoreCase))
            {
                InstanceSheet instanceUf = _insSheets.FirstOrDefault(x => x.Value.Name.Equals("TestInst_UF")).Value;
                if (instanceUf == null)
                {
                    return null;
                }

                InstanceRow instanceRow = instanceUf.Rows.FirstOrDefault(x => x.TestName.Equals(sequence.SubFlowName));
                if (instanceRow != null)
                {
                    flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = sequence.SubFlowName, Enable = sequence.EnableWord, FailAction = sequence.FailFlag };
                    AddRowIntoFlow(flowRow, ref mainFlow, sequence);
                    foreach (FlowRow bintableRow in CreateBinTable(sequence.FailFlag.Split(',').ToList(), "", sequence.BintableEnableWord))
                    {
                        AddRowIntoFlow(bintableRow, ref mainFlow, sequence);
                    }
                }
            }
            else
            {
                ConvertToRealFlowName(sequence, ref mainFlow, avalibleParts);
            }
            return flowRow;
        }

        private void AddTestNameWidthRow(int startIdx, string currentModule, string module, ref MainFlow flow)
        {
            if (startIdx == flow.Rows.Count || currentModule == module || TestPlanStatic.TestNameWidthTable == null)
            {
                return;
            }

            TestNameWidthRow target = TestPlanStatic.TestNameWidthTable.Rows.Find(x => x.Module.Equals(module));
            if (target == null)
            {
                return;
            }

            string instanceName = $"TestNameWidth_{module}";
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("DataloggerCustomize", "basic", true);
            if (function.Type == ".NET" && !_commonInstanceSheet.Rows.Exists(x => x.TestName.Equals(instanceName)))
            {
                function.SetParamValue("ptrTestNameWidth", target.PtrTestNameWidt);
                function.SetParamValue("ptrPinWidth", target.PtrPinWidth);
                function.SetParamValue("ftrTestNameWidth", target.FtrTestNameWidt);
                function.SetParamValue("ftrPatternWidth", target.FtrPatternWidth);
                _commonInstanceSheet.AddRow(new InstanceRow { TestName = instanceName, ArgList = function.Parameters, Args = function.ArgList, VbtName = function.FullFunctionName, VbtType = function.Type });
            }
            flow.InsertRow(startIdx, new FlowRow { Opcode = OpCode.Test, Parameter = instanceName });
        }

        internal static List<FlowRow> SetDeviceCondition(FlowSequenceNew sequence, FlowRow row, string jobs)
        {
            string siteVar = sequence.SiteFlagPerSite;
            FlowRow flowRow = new FlowRow();
            List<FlowRow> ifFlowRows = new List<FlowRow>();

            if (!string.IsNullOrEmpty(siteVar))
            {
                if (siteVar.Contains("&&") || siteVar.Contains("||"))
                {
                    ifFlowRows.Add(FlowRow.GenIfCondition(siteVar, ""));
                    ifFlowRows.Add(row);
                    ifFlowRows.Add(FlowRow.GenEndIf(""));
                }
                else
                {
                    AddCondition(siteVar, ref row);
                    ifFlowRows.Add(row);
                }
            }
            else
            {
                ifFlowRows.Add(row);
            }

            ifFlowRows.ForEach(x => x.RowNum = sequence.RowNum);
            ifFlowRows.ForEach(x => x.Job = jobs);
            return ifFlowRows;
        }

        internal static void AddCondition(string siteVar, ref FlowRow flowRow)
        {
            flowRow.DeviceName = siteVar.TrimStart('!');
            if (Regex.IsMatch(siteVar, @"^site-var", RegexOptions.IgnoreCase))
            {
                string siteVarCondition = siteVar.Split(' ').First();
                string siteVarValue = siteVar.Split(' ')[1] + ' ' + siteVar.Split(' ')[2];
                flowRow.DeviceCondition = siteVarCondition;
                flowRow.DeviceName = siteVarValue;
            }
            else if (!siteVar.Trim().EndsWith("False", StringComparison.CurrentCultureIgnoreCase))
            {
                flowRow.DeviceCondition = siteVar.StartsWith("!") ? "Flag-false" : "Flag-true";
            }
            else
            {
                flowRow.DeviceCondition = siteVar.StartsWith("!") ? "Flag-true" : "Flag-false";
            }
        }

        private void AddVariable(ref MainFlow mainFlow, string jobName, List<string> allJobs = null)
        {
            Dictionary<string, List<FlowRow>> allJobVariableRows = new Dictionary<string, List<FlowRow>>();
            if (allJobs != null)
            {
                foreach (string job in allJobs)
                {
                    if (!allJobVariableRows.ContainsKey(job))
                    {
                        allJobVariableRows.Add(job, new List<FlowRow>());
                    }
                }
            }

            if (_variableInitTable != null)
            {
                foreach (VariableRow row in _variableInitTable.Rows)
                {
                    string opcode = row.Opcode;
                    string parameter = row.Parameter;
                    if (jobName.Equals("All", StringComparison.CurrentCultureIgnoreCase) && allJobs != null)
                    {
                        foreach (string job in allJobs)
                        {
                            bool existJob = row.JobValues.ContainsKey(job);
                            if (!existJob)
                            {
                                continue;
                            }

                            string value = row.JobValues[job];
                            if (opcode.Equals("create-assign-site-var", StringComparison.OrdinalIgnoreCase))
                            {
                                allJobVariableRows[job].Add(new FlowRow { Opcode = OpCode.CreateSiteVar, Parameter = parameter, Job = job });
                                allJobVariableRows[job].Add(new FlowRow { Opcode = OpCode.AssignSiteVar, Parameter = parameter + " " + value, Job = job });
                            }
                            else if (opcode.Equals("create-assign-integer", StringComparison.OrdinalIgnoreCase))
                            {
                                allJobVariableRows[job].Add(new FlowRow { Opcode = OpCode.CreateInteger, Parameter = parameter, Job = job });
                                allJobVariableRows[job].Add(new FlowRow { Opcode = OpCode.AssignInteger, Parameter = parameter + " " + value, Job = job });
                            }
                        }

                    }
                    else
                    {
                        bool existJob = row.JobValues.ContainsKey(jobName);
                        if (!existJob)
                        {
                            continue;
                        }

                        string value = row.JobValues[jobName];
                        if (opcode.Equals("create-assign-site-var", StringComparison.OrdinalIgnoreCase))
                        {
                            mainFlow.AddRow(new FlowRow { Opcode = OpCode.CreateSiteVar, Parameter = parameter });
                            mainFlow.AddRow(new FlowRow { Opcode = OpCode.AssignSiteVar, Parameter = parameter + " " + value });
                        }
                        else if (opcode.Equals("create-assign-integer", StringComparison.OrdinalIgnoreCase))
                        {
                            mainFlow.AddRow(new FlowRow { Opcode = OpCode.CreateInteger, Parameter = parameter });
                            mainFlow.AddRow(new FlowRow { Opcode = OpCode.AssignInteger, Parameter = parameter + " " + value });
                        }
                    }
                }
                if (allJobVariableRows.Any())
                {
                    foreach (KeyValuePair<string, List<FlowRow>> jobVariable in allJobVariableRows)
                    {
                        mainFlow.AddRows(jobVariable.Value);
                    }
                }
            }
        }

        private List<SubFlowSheet> GetSubFlowsBySequence(FlowSequenceNew sequence)
        {
            string sourceSheet = sequence.SheetName;
            string subFlowSheet = string.IsNullOrEmpty(sequence.SubFlowName) ? "" : sequence.SubFlowName;
            List<string> flowNames = GetRealFlowSheetNames(sourceSheet, subFlowSheet);
            return _subFlowSheets.Values.Where(x => flowNames.Contains(x.Name)).ToList();
        }

        private List<string> GetRealFlowSheetNames(string sourceSheet, string subFlowSheet)
        {
            SheetNameMappingTable targetSheet = _sheetNameMappingTables.Find(x => x.SourceSheet.Equals(sourceSheet, StringComparison.OrdinalIgnoreCase) && x.SubFlowSheet.Equals(subFlowSheet, StringComparison.OrdinalIgnoreCase));
            List<string> flowNames = new List<string>();
            if (targetSheet == null)
            {
                if (string.IsNullOrEmpty(subFlowSheet))
                {
                    if (_sheetNameMappingTables.Exists(x => x.SourceSheet.Equals(sourceSheet, StringComparison.OrdinalIgnoreCase)))
                    {
                        List<SheetNameMappingTable> foundSheets = _sheetNameMappingTables.FindAll(x => x.SourceSheet.Equals(sourceSheet, StringComparison.OrdinalIgnoreCase));
                        foreach (SheetNameMappingTable foundSheet in foundSheets)
                        {
                            flowNames.AddRange(foundSheet.SheetName);
                        }
                    }
                    else if (_sheetNameMappingTables.Exists(x => x.SubFlowSheet.Equals(sourceSheet, StringComparison.OrdinalIgnoreCase)))
                    {
                        flowNames.Add(sourceSheet);
                    }
                    else if (_sheetNameMappingTables.Exists(x => x.SheetName.Exists(y => RegexFlow.Replace(y, "").Equals(sourceSheet, StringComparison.OrdinalIgnoreCase))))
                    {
                        SheetNameMappingTable row = _sheetNameMappingTables.Find(x => x.SheetName.Exists(y => RegexFlow.Replace(y, "").Equals(sourceSheet, StringComparison.OrdinalIgnoreCase)));
                        if (row.SheetName.Count == 1)
                        {
                            flowNames.Add(row.SheetName[0]);
                        }
                    }
                    else if (_sheetNameMappingTables.Exists(x => x.SheetName.Exists(y => y.Equals(sourceSheet, StringComparison.OrdinalIgnoreCase))))
                    {
                        flowNames.Add(sourceSheet);
                    }
                }
            }
            else
            {
                if (sourceSheet.Equals(InitializeConst))
                {
                    flowNames = targetSheet.SheetName.OrderBy(x => x).ToList();
                }
                else
                {
                    flowNames = targetSheet.SheetName;
                }
            }
            return flowNames;
        }

        private void ConvertToRealFlowName(FlowSequenceNew sequence, ref MainFlow mainflow, List<string> avalibleParts, string jobs = "")
        {
            string sourceSheet = sequence.SheetName;
            string subFlowSheet = string.IsNullOrEmpty(sequence.SubFlowName) ? "" : sequence.SubFlowName;
            List<string> flowNames = GetRealFlowSheetNames(sourceSheet, subFlowSheet);
            string part = "";
            if (!string.IsNullOrEmpty(sequence.Part) && avalibleParts != null && !string.Join(",", avalibleParts).Equals(sequence.Part, StringComparison.OrdinalIgnoreCase))
            {
                part = sequence.Part;
            }

            foreach (string name in flowNames)
            {
                FlowRow flowRow = new FlowRow
                {
                    Opcode = OpCode.Call,
                    Parameter = name,
                    Enable = sequence.EnableWord,
                    Part = part,
                    FailAction = sequence.FailFlag
                };
                AddRowIntoFlow(flowRow, ref mainflow, sequence, jobs);
                SetTempMonFromFlowSequence(sequence, flowRow);

                if (name.Equals(ConFlowTableMainInitEnableWd))
                {
                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.ConInitializeModulesFromInstance, "", true);
                    if (function.Type == ".NET")
                    {
                        flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = FuncNameConst.ConInitializeModulesFromInstance };
                        AddRowIntoFlow(flowRow, ref mainflow, sequence, jobs);
                    }
                }
                foreach (FlowRow bintableRow in CreateBinTable(sequence.FailFlag.Split(',').ToList()))
                {
                    AddRowIntoFlow(bintableRow, ref mainflow, sequence, jobs);
                }
            }
        }

        private void AddRowIntoFlow(FlowRow row, ref MainFlow mainFlow, FlowSequenceNew sequence, string jobs = "")
        {
            if (!string.IsNullOrEmpty(sequence.Group) && _groupSubFlow.Exists(x => x.Name.Equals(sequence.Group)))
            {
                SubFlowSheet targetSubflow = _groupSubFlow.Find(x => x.Name.Equals(sequence.Group, StringComparison.OrdinalIgnoreCase));
                FlowRow targetRow = targetSubflow.Rows.Find(x =>
                    x.Parameter.Equals(row.Parameter, StringComparison.OrdinalIgnoreCase)
                    && x.RowNum == sequence.RowNum);
                string mappingJob = sequence.Job;
                if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.JobMappingDic.ContainsKey(mappingJob))
                {
                    mappingJob = TestPlanStatic.TestProgramDefSheet.JobMappingDic[mappingJob];
                }
                List<string> targetRowJobs = new List<string> { mappingJob };

                if (targetRow != null)
                {
                    targetRowJobs = targetRow.GetJobs();
                    if (!targetRowJobs.Exists(x => x.Equals(mappingJob)))
                    {
                        targetRowJobs.Add(mappingJob);
                    }

                    targetRow.Job = string.Join(",", targetRowJobs);
                }
                else
                {
                    row.Job = string.Join(",", targetRowJobs);

                    string mergingjobs = string.Join(",",
                        jobs.Split(',', StringSplitOptions.RemoveEmptyEntries)
                            .Select(x => x.Trim())
                            .Concat(targetRowJobs.Where(x => !string.IsNullOrWhiteSpace(x)))
                            .Distinct()
                    );

                    targetSubflow.AddRows(SetDeviceCondition(sequence, row, mergingjobs));
                }
            }
            else
            {
                mainFlow.AddRows(SetDeviceCondition(sequence, row, jobs));
            }

            sequence.Found = true;
        }

        private List<FlowRow> CreateBinTable(List<string> flags, string bintableName = "", string bintableEnableWord = "")
        {
            List<FlowRow> flowRows = new List<FlowRow>();
            BinTableSheet binTableSheet = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
            foreach (string flag in flags)
            {
                if (string.IsNullOrEmpty(flag))
                {
                    break;
                }

                string binName = !string.IsNullOrEmpty(bintableName) ? bintableName : Regex.Replace(flag, "^F", "Bin");
                flowRows.Add(new FlowRow { Opcode = OpCode.BinTable, Parameter = binName, Enable = bintableEnableWord });
                if (binTableSheet.Rows.Exists(x => x.Name.Equals(binName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                BinTableRow binTableRow = Regex.IsMatch(flag, @"\!?\w+((\&\&)|(\|\|))\!?\w+") ? CreateBintableWithLogical(flag, binName) : new BinTableRow
                {
                    Name = binName,
                    ItemList = flag,
                    Items = new List<string> { "T" },
                    Op = "AND"
                };

                string[] binNameSegments = binName.Split('_');
                string category1 = binNameSegments.Length > 1 ? binNameSegments[1] : "";
                string category2 = binNameSegments.Length > 2 ? binNameSegments[2] : "";
                BinNumResult binInfo = BinNumberSingleton.Instance.GetBinInfo("FlowMain", category1, category2, binTableRow);

                binTableRow.Sort = binInfo.SoftBin.ToString("G15");
                binTableRow.Bin = binInfo.BinNumInfo.HardBin.ToString("G15");
                binTableRow.Result = binInfo.BinNumInfo.Status;
                binTableSheet.AddRow(binTableRow);
            }
            return flowRows;
        }

        private BinTableRow CreateBintableWithLogical(string flags, string binName)
        {
            Dictionary<string, string> flagsDic = flags.Split(new[] { '&', '|' }, StringSplitOptions.RemoveEmptyEntries).ToDictionary(flag => flag.Replace("!", ""), flag => flag.StartsWith("!") ? "F" : "T");
            BinTableRow bintableRow = new BinTableRow
            {
                Name = binName,
                ItemList = string.Join(",", flagsDic.Keys),
                Items = flagsDic.Values.ToList(),
                Op = flags.Contains("&&") ? "AND" : "OR"
            };
            return bintableRow;
        }

        private void GenGroupSubFlowSheets(List<MainFlowBase> mainFlow)
        {
            Dictionary<string, List<FlowSequence>> subMainFlowGroups = mainFlow.SelectMany(p => p.Sequences).Where(y => !string.IsNullOrEmpty(y.GroupSheetName))
                .GroupBy(p => p.GroupSheetName).ToDictionary(p => p.Key, p => p.ToList());
            foreach (KeyValuePair<string, List<FlowSequence>> subflow in subMainFlowGroups)
            {
                Dictionary<string, List<FlowSequence>> subflowGroups = subflow.Value.GroupBy(p => p.Job).ToDictionary(p => p.Key, p => p.ToList());
                SubFlowSheet flowSheet = new SubFlowSheet(subflow.Key) { JobNames = subflowGroups.SelectMany(x => x.Value).Select(x => x.Job).Distinct().ToList() };

                foreach (KeyValuePair<string, List<FlowSequence>> subFlowGroup in subflowGroups)
                {
                    foreach (FlowSequence jobItem in subFlowGroup.Value)
                    {
                        FlowRow flow = new FlowRow
                        {
                            Opcode = _subFlowSheets.Any(p => p.Value.Name.Equals
                                (subFlowGroup.Key, StringComparison.OrdinalIgnoreCase))
                                ? "call"
                                : "nop",
                            Parameter = jobItem.SubFlowName,
                            Job = subFlowGroup.Key
                        };
                        flowSheet.AddRow(flow);
                    }
                }
                flowSheet.AddReturnRow();
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirMain, flowSheet);
            }
        }

        private MainFlow GenMainFlowSheet(MainFlowBase flowBase)
        {
            string strSheetName = flowBase.MainFlowName;
            List<MainFlowData> subFlowListUse = new List<MainFlowData>();
            List<MainFlowData> subFlowListNonUse = new List<MainFlowData>();
            GetSubFlowListAndSubFlowListNoUse(ref subFlowListUse, ref subFlowListNonUse);

            MainFlow flowSheet = new MainFlow(strSheetName) { JobNames = new List<string> { flowBase.JobName } };

            #region Start Flow
            flowSheet.AddRow(GenAlarmRow());
            flowSheet.AddRow(GenErrorRow());

            if (LocalSpecs.Options.Device == EnumDevice.LCD && BlockStatus.GetAutomationBlockStatus(BlockStatus.Otp).Down) //TODO: (Jake) There is no rules...just customize with it
            {
                flowSheet.AddRows(GenLcdFlowPrintAndOthers());
            }

            flowSheet.AddRows(GenSrcCodeIndx());

            if (SelSramPatternSingleton.GetInstance().SelSramDatas.Count != 0)
            {
                flowSheet.AddRows(GenSelSramVar());
            }

            flowSheet.AddRows(GenInitVarFlow());

            flowSheet.AddRows(GenEfuseFlowVarInit());

            flowSheet.AddRow(GenFlowTableMainInitEnableWd(subFlowListUse));

            flowSheet.AddRow(GenFlowTableMainInitFlows(subFlowListUse));

            flowSheet.AddRow(GenFlowTableWriteSpiRomMain(flowBase, subFlowListUse));
            #endregion

            GetMainFlow(flowBase, subFlowListUse, ref flowSheet);

            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Efuse).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Efuse))
            {
                GenEfuseFlowToMain(subFlowListNonUse, ref flowSheet);
            }

            #region End Flow

            #region Non Use Flow

            foreach (FlowRow flowRow in flowSheet.Rows)
            {
                if (flowRow != null && !string.IsNullOrEmpty(flowRow.Parameter))
                {
                    SubFlowListNonUseRemove(subFlowListNonUse, flowRow.Parameter);
                }
            }

            SubFlowListNonUseRemove(subFlowListNonUse, ConSpirom32M);

            RemoveSubFlowListNonUse(subFlowListNonUse, flowBase);

            flowSheet.AddRows(GenNonUseFlow(subFlowListNonUse, flowSheet));
            #endregion

            flowSheet.AddRow(new FlowRow { Opcode = OpCode.LimitsAll });

            flowSheet.AddRow(GenSetDevice());
            #endregion

            return flowSheet;
        }

        private void SubFlowListNonUseRemove(List<MainFlowData> subFlowListNonUse, string name)
        {
            for (int i = subFlowListNonUse.Count - 1; i >= 0; i--)
            {
                if (subFlowListNonUse[i].FullSheetName.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                {
                    subFlowListNonUse.RemoveAt(i);
                }
            }
        }

        private List<FlowRow> GenNonUseFlow(List<MainFlowData> subFlowListNonUse, SubFlowSheet flowSheet)
        {
            List<FlowRow> flowRows = new List<FlowRow>();
            List<string> usedFlowSheets = flowSheet.GetUsedFlowSheets(_subFlowSheets.Values.ToList());
            foreach (MainFlowData nonUseSheet in subFlowListNonUse)
            {
                if (usedFlowSheets.Any(x => x.Equals(nonUseSheet.FullSheetName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    continue;
                }

                // flowSheet is main flow and have only on job
                if (flowSheet.JobNames != null &&
                    flowSheet.JobNames.Count == 1 &&
                    nonUseSheet.JobNames != null &&
                    nonUseSheet.JobNames.Any() &&
                    !nonUseSheet.JobNames.Exists(x => x.Equals(flowSheet.JobNames.First(), StringComparison.CurrentCultureIgnoreCase)))
                {
                    continue;
                }

                string env = "";
                string strEnable = "";
                if (Regex.IsMatch(nonUseSheet.FullSheetName, @"\dDCZ", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(nonUseSheet.FullSheetName, "ShiftInVs", RegexOptions.IgnoreCase))
                {
                    strEnable = "CZ_ENABLE";
                    env = "X";
                }
                if (nonUseSheet.FullSheetName.Equals("Flow_Post_IDS_EQN_Voltage"))
                {
                    int flowIdx = flowSheet.Rows.FindLastIndex(x => x.Parameter.ContainsIgnoreCase("FLOW_DCTEST_IDS"));
                    flowSheet.Rows.Insert(flowIdx + 1, GenMainFlowRow(nonUseSheet.FullSheetName, "", ConCommentDidnTSpecifyInFlowMainSheetTestplan, OpCode.Call, "Enable_ApplyVoltageFromBinCut"));
                }
                else
                {
                    flowRows.Add(GenMainFlowRow(nonUseSheet.FullSheetName, env, ConCommentDidnTSpecifyInFlowMainSheetTestplan, OpCode.Call, strEnable));
                }
            }
            return flowRows;
        }

        private FlowRow GenSetDevice()
        {
            FlowRow row = new FlowRow
            {
                Opcode = ConSetDevice,
                BinPass = "1",
                SortPass = "1",
                Result = "Pass"
            };
            return row;
        }

        private void GetSubFlowListAndSubFlowListNoUse(ref List<MainFlowData> subFlowListUse, ref List<MainFlowData> subFlowListNonUse)
        {
            foreach (KeyValuePair<string, SubFlowSheet> subFlow in _subFlowSheets)
            {
                MainFlowData mainFlowData = new MainFlowData();
                if (subFlow.Value.JobNames != null && subFlow.Value.JobNames.Any())
                {
                    mainFlowData.JobNames = subFlow.Value.JobNames;
                    mainFlowData.FullSheetName = subFlow.Value.Name;
                }
                else
                {
                    mainFlowData.FullSheetName = subFlow.Value.Name;
                }

                subFlowListUse.Add(mainFlowData);
                subFlowListNonUse.Add(mainFlowData);
            }
        }

        private void RemoveSubFlowListNonUse(List<MainFlowData> subFlowListNonUse, MainFlowBase flowBase)
        {
            foreach (SubFlowSheet subFlow in _subFlowSheets.Values)
            {
                if (subFlow.Name.StartsWith(ConFlowTablenWire, StringComparison.OrdinalIgnoreCase))
                {
                    SubFlowListNonUseRemove(subFlowListNonUse, subFlow.Name);
                }

                if (subFlow.Name.EndsWith(SuffixShmooFlowSheet, StringComparison.OrdinalIgnoreCase))
                {
                    SubFlowListNonUseRemove(subFlowListNonUse, subFlow.Name);
                }
            }

            if (flowBase.Sequences.Any(a => Regex.IsMatch(a.SubFlowName, "Flow_eFuse_IDS_Bincut_PreCheck|Flow_eFuse_CFG_UDR_MON_Read", RegexOptions.IgnoreCase)))
            {
                SubFlowListNonUseRemove(subFlowListNonUse, "Flow_eFuse_CFG_UDR_Read");
            }
        }

        private void GetMainFlow(MainFlowBase flowBase, List<MainFlowData> subFlowList, ref MainFlow flowSheet)
        {
            List<string> usedGroup = new List<string>();
            foreach (FlowSequence sequence in flowBase.Sequences)
            {
                string mainFlowName = sequence.SubFlowName;
                if (IsEfuseSubFlow(mainFlowName) && subFlowList.Exists(x => x.FullSheetName.Equals(mainFlowName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(sequence.GroupSheetName))
                {
                    mainFlowName = sequence.GroupSheetName;
                    if (usedGroup.Contains(sequence.GroupSheetName))
                    {
                        continue;
                    }

                    usedGroup.Add(mainFlowName);
                }
                string env = GenerateEnv(sequence, subFlowList);
                string comment = GenerateComment(mainFlowName, subFlowList);
                if (mainFlowName.Equals("Flow_Post_IDS_EQN_Voltage", StringComparison.OrdinalIgnoreCase))
                {
                    flowSheet.AddRow(GenMainFlowRow(mainFlowName, env, comment, OpCode.Call, "Enable_ApplyVoltageFromBinCut"));
                }
                else if (mainFlowName.ToUpper().StartsWith("CONCURRENT"))
                {
                    ConcurrentFlowSheet concurrentFlow = TestPlanStatic.ConcurrentFlowSheet;
                    ConcurrentFlowSheetRow concurrentRow =
                        concurrentFlow.Rows.Find(
                            x => x.SequenceName.Equals(mainFlowName.Split(new[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries).Last().Trim()));
                    foreach (string subFlow in concurrentRow.Subflows)
                    {
                        flowSheet.AddRow(GenMainFlowRow(subFlow, "", comment, OpCode.Call, "Serial_case"));
                    }
                    flowSheet.AddRow(GenMainFlowRow(concurrentRow.SequenceName, "", comment, OpCode.Concurrent, "sequence"));
                    flowSheet.HasConcurrent = true;
                }
                else
                {
                    flowSheet.AddRow(GenMainFlowRow(mainFlowName, env, comment, OpCode.Call));
                }
            }
        }

        private FlowRow GenFlowTableWriteSpiRomMain(MainFlowBase flowBase, List<MainFlowData> subFlowList)
        {
            // Create SPIROM if flow contains either IDS or SPI.
            if (IsContain(ConIds, flowBase.Sequences.Select(p => p.SubFlowName).ToList()) ||
                IsContain(ConSpi, flowBase.Sequences.Select(p => p.SubFlowName).ToList()))
            {
                string env = GenerateEnv(new FlowSequence(ConSpiromMain), subFlowList);
                string comment = GenerateComment(ConSpiromMain);
                return GenMainFlowRow(ConSpiromMain, env, comment, OpCode.Call);
            }
            return null;
        }

        private FlowRow GenFlowTableMainInitFlows(List<MainFlowData> subFlowList)
        {
            string env = GenerateEnv(new FlowSequence(ConFlowTableMainInitFlows), subFlowList);
            return GenMainFlowRow(ConFlowTableMainInitFlows, env, "", OpCode.Call);
        }

        private FlowRow GenFlowTableMainInitEnableWd(List<MainFlowData> subFlowList)
        {
            string env = GenerateEnv(new FlowSequence(ConFlowTableMainInitEnableWd), subFlowList);
            return GenMainFlowRow(ConFlowTableMainInitEnableWd, env, "", OpCode.Call);
        }

        private List<FlowRow> GenSelSramVar()
        {
            List<FlowRow> flowRows = new List<FlowRow>();
            FlowRow flowRow = new FlowRow { Opcode = "create-site-var", Parameter = "SelSrm_LP_Var" };
            flowRows.Add(flowRow);

            flowRow = new FlowRow { Opcode = "create-site-var", Parameter = string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) ? "LPCount_End" : "SEL_PAT_END" };
            flowRows.Add(flowRow);

            flowRow = new FlowRow { Opcode = "create-site-var", Parameter = "SEL_CASE_End" };
            flowRows.Add(flowRow);

            flowRow = new FlowRow { Opcode = "create-site-var", Parameter = "SEL_CASE_CNT" };
            flowRows.Add(flowRow);
            return flowRows;
        }

        private List<FlowRow> GenSrcCodeIndx()
        {
            List<FlowRow> flowRows = new List<FlowRow>();
            foreach (string usedinteger in HardIpStatic.FlowUsedInteger)
            {
                FlowRow srcCodeIndxRow = new FlowRow
                {
                    Opcode = "create-integer",
                    Parameter = usedinteger
                };
                flowRows.Add(srcCodeIndxRow);
            }
            return flowRows;
        }

        private FlowRow GenErrorRow()
        {
            BinNumResult binInfo = BinNumberSingleton.Instance.GetBinInfo("set-error-bin", "", "");
            FlowRow erroRow = new FlowRow
            {
                Opcode = "set-error-bin",
                BinFail = binInfo.Found ? binInfo.BinNumInfo.HardBin.ToString("G15") : "9",
                SortFail = binInfo.Found ? binInfo.SoftBin.ToString("G15") : "999"
            };
            return erroRow;
        }

        private FlowRow GetInitTestTimeUtilityRow()
        {
            if (_insSheets.Count != 0)
            {
                if (_commonInstanceSheet != null && !_commonInstanceSheet.Rows.Exists(x => x.TestName.Equals("InitTestTimeUtility")))
                {
                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName("InitTestTimeUtility", "basic", true);
                    if (function.Type == ".NET")
                    {
                        _commonInstanceSheet.Rows.Add(new InstanceRow { TestName = "InitTestTimeUtility", VbtType = function.Type, VbtName = function.FullFunctionName, ArgList = function.Parameters, Args = function.ArgList });
                        return new FlowRow { Opcode = OpCode.Test, Parameter = "InitTestTimeUtility" };
                    }
                    return null;
                }

                if (_commonInstanceSheet != null && _commonInstanceSheet.Rows.Exists(x => x.TestName.Equals("InitTestTimeUtility")))
                {
                    return new FlowRow { Opcode = OpCode.Test, Parameter = "InitTestTimeUtility" };
                }
            }
            return null;
        }

        private FlowRow GenAlarmRow()
        {
            BinNumResult binInfo = BinNumberSingleton.Instance.GetBinInfo("set-alarm-bin", "", "");
            FlowRow erroRow = new FlowRow
            {
                Opcode = "set-alarm-bin",
                BinFail = binInfo.Found ? binInfo.BinNumInfo.HardBin.ToString("G15") : "9",
                SortFail = binInfo.Found ? binInfo.SoftBin.ToString("G15") : "998"
            };
            return erroRow;
        }

        internal string GenerateEnv(FlowSequence flowSequence, List<MainFlowData> subFlowList)
        {
            string strResult = "";
            if (!string.IsNullOrEmpty(flowSequence.GroupSheetName))
            {
                return "";
            }

            //nop all efuse item in TP
            if (IsEfuseSubFlow(flowSequence.SubFlowName))
            {
                return "X";
            }

            foreach (MainFlowData flowSheet in subFlowList)
            {
                string igSheet = flowSheet.FullSheetName.ToUpper().Replace("_", "");
                string testPlanSheet = flowSequence.SubFlowName.ToUpper().Replace("_", "");
                if (igSheet.Contains(testPlanSheet))
                {
                    strResult = "";
                    break;
                }
                strResult = "X";
            }

            //if the subFlow is Td Flow,
            //Make a judgement if Flow list has BinCut Flow
            //if has BinCut unable Td Flow
            if (Regex.IsMatch(flowSequence.SubFlowName, ConKeywordTd, RegexOptions.IgnoreCase))
            {
                if (subFlowList.Exists(x => x.FullSheetName.IndexOf(ConVddBiningFlow, StringComparison.CurrentCultureIgnoreCase) != -1))
                {
                    strResult = "X";
                }
            }
            return strResult;
        }

        internal string GetBlockDisabledComment(string block)
        {
            string strResult = "";
            if (!BlockStatus.GetAutomationBlockStatus(block).Enable)
            {
                strResult = ConCommentUnable;
            }
            if (BlockStatus.GetAutomationBlockStatus(block).Enable && !BlockStatus.GetAutomationBlockStatus(block).Down)
            {
                strResult = ConCommentUnableFromUi;
            }
            return strResult;
        }

        internal string GenerateComment(string subFlow, List<MainFlowData> subFlowList = null)
        {
            const string strConti = "^Flow_DC_Conti$";
            const string strScan = @"^Flow_\w*Scan.*$";
            const string strMbist = @"^Flow_\w*Mbist.*$";
            const string strHardIp = "^Flow_HARDIP.*$";
            const string strSpi = "^Flow_SPI$";
            const string strBinCut = ConVddBiningFlow;
            const string strEfuse = "^Flow_eFuse.*$";
            const string strSpiRom = "^Flow_Table_Write_SPIROM_Main$";
            const string strIds = "Flow_DCTEST_IDS";
            const string strGpio = "^Flow_DCTEST_GPIO$";
            const string strFailSafe = "^DCTEST_FailSafe$";
            const string strAllBlankChk = "^Flow_eFuse_ECID_All_Blank_Check$";
            string strResult = "";

            //Basic
            if (Regex.IsMatch(subFlow, strConti, RegexOptions.IgnoreCase) ||
                Regex.IsMatch(subFlow, strSpiRom, RegexOptions.IgnoreCase))
            {
                strResult = GetBlockDisabledComment(BlockStatus.Basic);
            }

            //Scan
            else if (Regex.IsMatch(subFlow, strScan, RegexOptions.IgnoreCase))
            {
                strResult = GetBlockDisabledComment(BlockStatus.Scan);
                if (strResult.Length == 0 && Regex.IsMatch(subFlow, ConKeywordTd, RegexOptions.IgnoreCase) &&
                    BlockStatus.GetAutomationBlockStatus(BlockStatus.BinCut).Enable && BlockStatus.GetAutomationBlockStatus(BlockStatus.BinCut).Down)
                {
                    strResult = ConCommentUnableByTool + " for VddBining";
                }
            }
            //Mbist
            else if (Regex.IsMatch(subFlow, strMbist, RegexOptions.IgnoreCase))
            {
                strResult = GetBlockDisabledComment(BlockStatus.Mbist);
            }
            //Hard IP
            else if (Regex.IsMatch(subFlow, strGpio, RegexOptions.IgnoreCase) ||
                Regex.IsMatch(subFlow, strFailSafe, RegexOptions.IgnoreCase) ||
                Regex.IsMatch(subFlow, strIds, RegexOptions.IgnoreCase) ||
                Regex.IsMatch(subFlow, strHardIp, RegexOptions.IgnoreCase))
            {
                strResult = GetBlockDisabledComment(BlockStatus.HardIp);
            }
            //SPI
            else if (Regex.IsMatch(subFlow, strSpi, RegexOptions.IgnoreCase))
            {
                strResult = GetBlockDisabledComment(BlockStatus.Rtos);
            }
            //BinCut
            else if (Regex.IsMatch(subFlow, strBinCut, RegexOptions.IgnoreCase))
            {
                strResult = GetBlockDisabledComment(BlockStatus.BinCut);
            }
            //Special EVS+eFuse
            else if (Regex.IsMatch(subFlow, strAllBlankChk, RegexOptions.IgnoreCase))
            {
                strResult = GetBlockDisabledComment(BlockStatus.Efuse);
                if (subFlowList != null)
                {
                    if (!subFlowList.Exists(x => x.FullSheetName.Equals("Flow_EVS", StringComparison.CurrentCultureIgnoreCase)))
                    {
                        strResult = strResult.Length > 0 ? strResult + ", " : strResult;
                        strResult += "Redudant item due to subflow Flow_EVS didn't exist";
                    }
                }
            }
            //Efuse
            else if (Regex.IsMatch(subFlow, strEfuse, RegexOptions.IgnoreCase))
            {
                strResult = GetBlockDisabledComment(BlockStatus.Efuse);
            }

            if (subFlow.StartsWith("Flow_", StringComparison.CurrentCultureIgnoreCase) &&
                !_subFlowSheets.Values.Select(x => x.Name).
               Any(x => x.Equals(subFlow, StringComparison.CurrentCultureIgnoreCase)))
            {
                strResult = ConCommentMissing;
            }

            return strResult;
        }

        internal bool IsContain(string str, List<string> strList)
        {
            bool lBResult = false;
            foreach (string s in strList)
            {
                if (s.ContainsIgnoreCase(str.ToUpper()))
                {
                    lBResult = true;
                }
            }
            return lBResult;
        }

        internal string ModifyMainFlowParameter(string flowSheet)
        {
            if (Regex.IsMatch(flowSheet, "^Flow_HARDIP.*$", RegexOptions.IgnoreCase) && !Regex.IsMatch(flowSheet, "init", RegexOptions.IgnoreCase)
                && Regex.IsMatch(flowSheet, @"^FLOW_\w+_\w+_\w+", RegexOptions.IgnoreCase) && !Regex.IsMatch(flowSheet, @"_CZ\d?$", RegexOptions.IgnoreCase))
            {

                string header = Regex.Match(flowSheet, "(?<header>^[a-zA-Z]+_[a-zA-Z]+_)").Groups["header"].ToString();
                string item = Regex.Match(flowSheet, @"^[a-zA-Z]+_[a-zA-Z]+_(?<name>\w+)").Groups["name"].ToString().Replace("_", "");
                string flowName = header + item;
                return flowName.ToUpper();
            }
            return flowSheet;
        }

        internal FlowRow GenMainFlowRow(string parameter, string env, string comment, string opcode, string enable = "")
        {
            FlowRow flowRow = new FlowRow
            {
                Opcode = env.Equals("") ? opcode : "Nop",
                Env = env,
                Enable = enable,
                Comment = comment,
                Parameter = ModifyMainFlowParameter(parameter)
            };
            return flowRow;
        }

        private List<FlowRow> GenInitVarFlow()
        {
            List<FlowRow> flowRows = new List<FlowRow>();
            SubFlowSheet initFlowSheet = null;
            const string strPattern = "^Flow_Table_Main_Init_Var*";
            foreach (KeyValuePair<string, SubFlowSheet> flow in _subFlowSheets)
            {
                if (Regex.IsMatch(flow.Value.Name, strPattern, RegexOptions.IgnoreCase))
                {
                    initFlowSheet = flow.Value;
                    break;
                }
            }

            if (initFlowSheet != null)
            {
                foreach (FlowRow row in initFlowSheet.Rows)
                {
                    if (row.Opcode.ToLower() == "return" ||
                        row.Opcode.ToLower() == "set-error-bin" ||
                        row.Opcode.ToLower() == "call")
                    {
                        continue;
                    }

                    flowRows.Add(row);
                }
            }
            return flowRows;
        }

        private List<FlowRow> GenEfuseFlowVarInit()
        {
            List<FlowRow> flowRows = LocalSpecs.EfuseFlowUsedInteger.Select(usedinteger => new FlowRow
            {
                Opcode = OpCode.CreateSiteVar,
                Parameter = usedinteger,
                ColumnA = EFuseConst.Efuse
            }).ToList();
            flowRows.AddRange(LocalSpecs.EfuseFlowUsedInteger.Select(usedinteger => new FlowRow
            {
                Opcode = OpCode.AssignSiteVar,
                Parameter = usedinteger + " 0",
                ColumnA = EFuseConst.Efuse
            }));

            return flowRows;
        }

        private void EnsureEfuseBankReadInstances(InstanceSheet currInstSheet)
        {
            if (currInstSheet == null)
            {
                return;
            }

            if (!currInstSheet.Rows.Exists(x => x.TestName.Equals(EFuseConst.PseudoFuseReadItem)))
            {
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(EFuseConst.PseudoFuseReadItem, "efuse", true);
                if (function.Type == ".NET")
                {
                    currInstSheet.AddRow(new EfuseGenerateInstance().AddCsharpInstanceItem(EFuseConst.PseudoFuseReadItem, EFuseConst.PseudoFuseReadItem));
                }
                else
                {
                    function = TestProgram.VbtFunctionLib.GetFunctionByName(EFuseConst.PseudoFuseReadItem, "efuse");
                    currInstSheet.AddRow(new EfuseGenerateInstance().AddCsharpInstanceItem(EFuseConst.PseudoFuseReadItem, EFuseConst.PseudoFuseReadItem));
                }
            }
            if (!currInstSheet.Rows.Exists(x => x.TestName.Equals(EFuseConst.BkmCp)))
            {
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Set_BKM", "efuse", true);
                if (function.Type == ".NET")
                {
                    currInstSheet.AddRow(new EfuseGenerateInstance().AddCsharpInstanceItem(EFuseConst.BkmCp, "Set_BKM"));
                }
                else
                {
                    function = TestProgram.VbtFunctionLib.GetFunctionByName("BKM_Update", "efuse");
                    currInstSheet.AddRow(new EfuseGenerateInstance().AddInstanceItem(EFuseConst.BkmCp, "BKM_Update"));
                }
            }
        }

        private void EnsureEfuseBankReadBinTables(BinTableSheet currBinTableSheet)
        {
            if (currBinTableSheet == null)
            {
                return;
            }

            if (!currBinTableSheet.Rows.Exists(x => x.Name.Equals("Bin_EFUSE_ReadFromFile")))
            {
                currBinTableSheet.AddRow(new EfuseBinTableWriter().GenBinTableRow("Bin_EFUSE_ReadFromFile"));
            }

            if (!currBinTableSheet.Rows.Exists(x => x.Name.Equals("Bin_EFUSE_BKM_Fail")))
            {
                currBinTableSheet.AddRow(new EfuseBinTableWriter().GenBinTableRow("Bin_EFUSE_BKM_Fail"));
            }

            if (!currBinTableSheet.Rows.Exists(x => x.Name.Equals("Bin_EFUSE_GoodBin_Retest")))
            {
                currBinTableSheet.AddRow(new EfuseBinTableWriter().GenBinTableRow("Bin_EFUSE_GoodBin_Retest"));
            }
        }

        private void GenEfuseFlowToMain(List<MainFlowData> subFlowListNonUse, ref MainFlow flowSheet)
        {
            int index = flowSheet.Rows.FindIndex(x => x.Parameter.StartsWith("Flow_HARDIP_JTAG", StringComparison.CurrentCultureIgnoreCase));
            if (index == -1)
            {
                index = flowSheet.Rows.FindIndex(x => x.Parameter.StartsWith("Flow_DC_Conti", StringComparison.CurrentCultureIgnoreCase));
            }

            MainFlowData bankReadFlow = subFlowListNonUse.Find(x => x.FullSheetName.Equals(Combination.CombineByUnderLine(EFuseConst.FlowSheetName, EFuseConst.BankRead)));
            MainFlowData evsAllBc = subFlowListNonUse.Find(x => x.FullSheetName.Equals(Combination.CombineByUnderLine(EFuseConst.FlowSheetName, EFuseConst.EvsAllBc)));
            InstanceSheet currInstSheet = _insSheets.FirstOrDefault(x => x.Value.Name.Equals(EFuseConst.TestInstSheet)).Value;
            BinTableSheet currBinTableSheet = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
            MainFlowData ecidDeid = subFlowListNonUse.Find(
                    x =>
                        x.FullSheetName.Equals(
                            Combination.CombineByUnderLine(new List<string>
                            {
                                EFuseConst.FlowSheetName,
                                BankType.Ecid,
                                EFuseConst.Deid
                            })));
            MainFlowData ecid = subFlowListNonUse.Find(
                    x =>
                        x.FullSheetName.Equals(
                            Combination.CombineByUnderLine(new List<string> { EFuseConst.FlowSheetName, BankType.Ecid })));

            if (index != -1)
            {
                index++;
            }

            if (bankReadFlow != null)
            {
                flowSheet.InsertRow(index, GenNewItemToMainFlow(bankReadFlow.FullSheetName, "!CP1", OpCode.Call, "!Pgm2File", EFuseConst.Efuse));
                flowSheet.InsertRow(++index, GenNewItemToMainFlow(EFuseConst.BkmCp, "CP1,CP2", OpCode.Test, "!Disable_eFuse_BKM", EFuseConst.Efuse,
                        "F_BKM_Fail"));
                flowSheet.InsertRow(++index, GenNewItemToMainFlow("Bin_EFUSE_BKM_Fail", "CP1,CP2", OpCode.BinTable, "!Disable_eFuse_BKM",
                        EFuseConst.Efuse));
                flowSheet.InsertRow(++index, GenNewItemToMainFlow(EFuseConst.PseudoFuseReadItem, "!CP1", OpCode.Test, "Pgm2File", EFuseConst.Efuse,
                         "F_ReadFromFile"));
                flowSheet.InsertRow(++index, GenNewItemToMainFlow("Bin_EFUSE_ReadFromFile", "!CP1", OpCode.BinTable, "Pgm2File", EFuseConst.Efuse));
                EnsureEfuseBankReadInstances(currInstSheet);

                EnsureEfuseBankReadBinTables(currBinTableSheet);

                if (evsAllBc != null)
                {
                    flowSheet.InsertRow(++index, GenNewItemToMainFlow(evsAllBc.FullSheetName, "CP1,CP2", OpCode.Nop, "!Pgm2File", EFuseConst.Efuse));
                    subFlowListNonUse.Remove(evsAllBc);
                }

                if (ecidDeid != null)
                {
                    flowSheet.InsertRow(++index, GenNewItemToMainFlow(ecidDeid.FullSheetName, "CP1", OpCode.Nop, "!Pgm2File&&!Char_CP", EFuseConst.Efuse));
                    subFlowListNonUse.Remove(ecidDeid);
                }
                if (ecid != null)
                {
                    flowSheet.InsertRow(++index, GenNewItemToMainFlow(ecid.FullSheetName, "", OpCode.Call, "!Pgm2File&&!Char_CP", EFuseConst.Efuse));
                    subFlowListNonUse.Remove(ecid);
                }
                subFlowListNonUse.Remove(bankReadFlow);
            }

            int tempIndex = index;
            index =
                flowSheet.Rows.FindIndex(
                    x => x.Parameter.StartsWith("Flow_DCTEST", StringComparison.CurrentCultureIgnoreCase));
            if (index != -1)
            {
                index++;
            }
            else
            {
                index = tempIndex;
            }

            MainFlowData cfgEarly =
                subFlowListNonUse.Find(
                    x =>
                        x.FullSheetName.Equals(
                            Combination.CombineByUnderLine(new List<string>
                            {
                                EFuseConst.FlowSheetName,
                                BankType.Cfg,
                                EFuseConst.Early
                            })));
            if (cfgEarly != null)
            {
                flowSheet.InsertRow(index,
                    GenNewItemToMainFlow(cfgEarly.FullSheetName, "CP1", OpCode.Call,
                        Combination.CombineByUnderLine(new List<string>
                        {
                            BankType.Cfg,
                            EFuseConst.Early,
                            EFuseConst.BankEnable
                        }), EFuseConst.Efuse));
                subFlowListNonUse.Remove(cfgEarly);
            }
            //put to the end
            MainFlowData preWrite =
                subFlowListNonUse.Find(
                    x =>
                        x.FullSheetName.Equals(Combination.CombineByUnderLine(EFuseConst.FlowSheetName, EFuseConst.PreWrite)));
            if (preWrite != null)
            {
                flowSheet.AddRow(GenNewItemToMainFlow(preWrite.FullSheetName, "", OpCode.Call, "", EFuseConst.Efuse));
                subFlowListNonUse.Remove(preWrite);
            }
            flowSheet.AddRow(GenNewItemToMainFlow("EFUSE_Check_Allbank_PatternFlag", "", OpCode.Test,
                "eFuse_Check_Allbank_Enable", EFuseConst.Efuse, "F_Check_PatternFlag"));
            flowSheet.AddRow(GenNewItemToMainFlow("Bin_EFUSE_Check_PatternFlag", "", OpCode.BinTable,
                "eFuse_Check_Allbank_Enable", EFuseConst.Efuse));
            if (currInstSheet != null)
            {
                if (!currInstSheet.Rows.Exists(x => x.TestName.Equals("EFUSE_Check_Allbank_PatternFlag")))
                {
                    currInstSheet.AddRow(new EfuseGenerateInstance().AddInstanceItem("EFUSE_Check_Allbank_PatternFlag",
                        "auto_eFuse_CheckPatternFlag"));
                }
            }

            if (currBinTableSheet != null)
            {
                if (!currBinTableSheet.Rows.Exists(x => x.Name.Equals("Bin_EFUSE_Check_PatternFlag")))
                {
                    currBinTableSheet.AddRow(new EfuseBinTableWriter().GenBinTableRow("Bin_EFUSE_Check_PatternFlag"));
                }

                if (!currBinTableSheet.Rows.Exists(x => x.Name.Equals("Bin_EFUSE_FlatCheck_Summary")))
                {
                    currBinTableSheet.AddRow(new EfuseBinTableWriter().GenBinTableRow("Bin_EFUSE_FlatCheck_Summary"));
                }
            }

            MainFlowData pseudoFuse = subFlowListNonUse.Find(x => x.FullSheetName.Equals(Combination.CombineByUnderLine(EFuseConst.FlowSheetName, EFuseConst.PseudoFuse)));
            if (pseudoFuse != null)
            {
                flowSheet.AddRow(GenNewItemToMainFlow(pseudoFuse.FullSheetName, "", OpCode.Call, "Pgm2File",
                    EFuseConst.Efuse));
                subFlowListNonUse.Remove(pseudoFuse);
            }
            List<string> efuseFlowOrder = new List<string>
            {
                //Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.Ecid),
                Combination.CombineByUnderLine(new List<string> {EFuseConst.FlowSheetName, BankType.Ecid, EFuseConst.Deid}),
                Combination.CombineByUnderLine(new List<string>
                {
                    EFuseConst.FlowSheetName,
                    BankType.Ecid,
                    EFuseConst.NonDeid
                }),
                Combination.CombineByUnderLine(new List<string> {EFuseConst.FlowSheetName, BankType.Cfg, EFuseConst.Early}),
                Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.Cfg),
                Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.UdrP),
                Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.UdrP0),
                Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.UdrP1),
                Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.UdrE),
                Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.UdrE0),
                Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.UdrE1),
                Combination.CombineByUnderLine(EFuseConst.FlowSheetName, BankType.Mon)
            };
            IOrderedEnumerable<MainFlowData> otherFuseSubFlows =
                subFlowListNonUse.FindAll(x => x.FullSheetName.StartsWith(EFuseConst.FlowSheetName))
                    .OrderBy(y => efuseFlowOrder.IndexOf(y.FullSheetName));
            foreach (MainFlowData fuseSubFlow in otherFuseSubFlows)
            {
                if (
                    fuseSubFlow.FullSheetName.Equals(
                        Combination.CombineByUnderLine(new List<string>
                        {
                            EFuseConst.FlowSheetName,
                            BankType.Ecid,
                            EFuseConst.NonDeid
                        })))
                {
                    flowSheet.AddRow(GenNewItemToMainFlow(fuseSubFlow.FullSheetName, "", OpCode.Nop, "!Pgm2File&&!Char_CP",
                        EFuseConst.Efuse));
                }
                else
                {
                    flowSheet.AddRow(GenNewItemToMainFlow(fuseSubFlow.FullSheetName, "", OpCode.Call, "!Pgm2File",
                    EFuseConst.Efuse));
                }
                subFlowListNonUse.Remove(fuseSubFlow);
            }
            flowSheet.AddRow(GenNewItemToMainFlow(EFuseConst.GoodBinRetestDetection, "", OpCode.Test,
                "!" + EFuseConst.BypReTestCheck, EFuseConst.Efuse, "F_EFUSE_GoodBin_Retest"));
            flowSheet.AddRow(GenNewItemToMainFlow("Bin_EFUSE_GoodBin_Retest", "", OpCode.BinTable,
                "!" + EFuseConst.BypReTestCheck, EFuseConst.Efuse));
        }

        internal FlowRow GenNewItemToMainFlow(string parameter, string job, string opcode, string enable = "", string columnA = "", string flag = "")
        {
            FlowRow flowRow = new FlowRow
            {
                Opcode = opcode,
                Enable = enable,
                Job = job,
                FailAction = string.IsNullOrEmpty(flag) ? "" : flag,
                ColumnA = columnA,
                Parameter = ModifyMainFlowParameter(parameter)
            };
            return flowRow;
        }

        internal bool IsEfuseSubFlow(string sheetName)
        {
            return sheetName.ContainsIgnoreCase("_efuse") || sheetName.ContainsIgnoreCase("_fuse");
        }

        private List<FlowRow> GenLcdFlowPrintAndOthers()
        {
            List<FlowRow> flowRows = new List<FlowRow>();
            //site var flow
            FlowRow siteFlowRow = new FlowRow { Opcode = OpCode.CreateSiteVar, Parameter = "Flag_OTPRead" };
            flowRows.Add(siteFlowRow);

            siteFlowRow = new FlowRow { Opcode = OpCode.AssignSiteVar, Parameter = "Flag_OTPRead -1" };
            flowRows.Add(siteFlowRow);

            siteFlowRow = new FlowRow { Opcode = OpCode.CreateSiteVar, Parameter = "RunTrim" };
            flowRows.Add(siteFlowRow);

            siteFlowRow = new FlowRow { Opcode = OpCode.AssignSiteVar, Parameter = "RunTrim -1" };
            flowRows.Add(siteFlowRow);

            //others
            FlowRow otpRows = new FlowRow { Opcode = OpCode.Call, Parameter = "FLOW_OTP_desc_bit_check" };
            flowRows.Add(otpRows);

            otpRows = new FlowRow { Opcode = OpCode.If, Parameter = "Flag_OTPRead=0" };
            flowRows.Add(otpRows);

            otpRows = new FlowRow
            {
                Opcode = OpCode.Test,
                Parameter = "OTP_Read_ALL_SetOTPdata_OneShot",
                TNum = "5000",
                FailAction = "F_OTP_REG_LAG1_idx"
            };
            flowRows.Add(otpRows);

            flowRows.Add(FlowRow.GenEndIf(""));

            return flowRows;
        }
    }
}
