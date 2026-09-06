using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.PostAction.BinCutELB;
using Automation.GenerateIgxl.PostAction.ConcurrentSequence;
using Automation.GenerateIgxl.PostAction.EndFlow;
using Automation.GenerateIgxl.PostAction.GenFlowTableMainInitEnableWd;
using Automation.GenerateIgxl.PostAction.GenIdsLeakagePatternSummary;
using Automation.GenerateIgxl.PostAction.GenJob;
using Automation.GenerateIgxl.PostAction.GenMainFlow;
using Automation.GenerateIgxl.PostAction.GenNonIgxlSheets;
using Automation.GenerateIgxl.PostAction.GenReferenceSheet;
using Automation.GenerateIgxl.PostAction.GenTestNumber;
using Automation.GenerateIgxl.PostAction.GenWireByPattern;
using Automation.GenerateIgxl.PostAction.InstCommon;
using Automation.GenerateIgxl.PostAction.ModifyChannelMap;
using Automation.GenerateIgxl.PostAction.PatSetProcessing;
using Automation.GenerateIgxl.PostAction.PinMapSort;
using Automation.GenerateIgxl.PostAction.PostCheck;
using Automation.GenerateIgxl.PostAction.SelSram;
using Automation.GenerateIgxl.PostAction.TempMon;
using Automation.GenerateIgxl.PostAction.UpdatePatternSetTimeDomain;
using Automation.GenerateIgxl.Scan.Harvest;
using Automation.PreCheck.AllParaData;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using ECIDSortAutogen;

using IgxlLib;
using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.Concurrent;
using TestPlanLib.Harvest;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.PostAction
{
    public class PostActionMain : WorkFlowBase<PostActionParaData>
    {
        private readonly List<string> _existedExtraSheet = new List<string>();

        public override bool PreCheckFlow(PostActionParaData paraData)
        {
            return true;
        }

        public override void WorkFlow(PostActionParaData postActionParaData)
        {
            postActionParaData = GetPostActionParas();
            try
            {
                HardIpInfos patInfo = LocalSpecs.HardIpInfos;
                if (VbtFunctionLibShared.UsedPatternList.Any() && LocalSpecs.SearchValidPatternRev)
                {
                    OptimizePatSetAll();
                }

                ErrorReportManager.RemoveNonUsedPattern(VbtFunctionLibShared.UsedPatternList);

                //Add extra N/C pins, Add extra DCVI, DCVS pins
                Response.Report("Modifying JobList Sheet ...", percentage: 10);
                var channelMapMain = new ModifyChannelMapMain();
                channelMapMain.WorkFlow();

                //Move pin group type is TimeDomain to the last item
                if (LocalSpecs.Options.Device == EnumDevice.AP)
                {
                    Response.Report("Modifying Pin Map Sheet ...", percentage: 15);
                    var pinMapProcess = new PinMapProcessMain();
                    pinMapProcess.WorkFlow();

                    var selSramMain = new SelSramMain();
                    selSramMain.WorkFlow();
                }

                //Gen CharSetup
                Response.Report("Generating CharSetup Sheet ...", percentage: 20);
                GenCharSetupSheets();

                // Add Flow_Post_IDS_EQN_Voltage subflow
                AddFlowPostIdsEqnVoltage();

                // Add Reference file
                Response.Report("Adding Reference file ...", percentage: 45);
                var genReferenceSheetMain = new GenReferenceSheetMain();
                genReferenceSheetMain.GenReferenceSheet();

                //Add nWire call before pattern
                Response.Report("Adding nWire before specified pattern ...", percentage: 60);
                var nWireByPattern = new NWireByPatternMain();
                nWireByPattern.WorkFlow(TestProgram.IgxlWorkBk.SubFlowSheets);

                //Add end of subflow
                Response.Report("Generating End Flow ...", percentage: 65);
                var endFlow = new MainEndFlow();
                endFlow.WorkFlow();

                //Add items from extraSheet
                GenItemsFromExtra();

                //Gen Harvest binTable
                GenHarvestBinTable();

                //Gen Main_Flow sheets
                Response.Report("Generating Main Flow ...", percentage: 75);
                var mainFlowGen = new MainFlowMain(TestPlanStatic.VariableInitTable, TestProgram.IgxlWorkBk.SubFlowSheets, TestProgram.IgxlWorkBk.InsSheets);
                mainFlowGen.WorkFlow(postActionParaData);

                if (EpWorkbook.TestPlanWorkbook.Worksheets["Concurrent_Flow"] != null)
                {
                    Response.Report("Generating Concurrent Sequence ...", percentage: 75);
                    ConcurrentFlowSheet concurrentSheet = TestPlanStatic.ConcurrentFlowSheet;
                    Dictionary<string, SubFlowSheet> flowSheets = TestProgram.IgxlWorkBk.SubFlowSheets;
                    Dictionary<string, InstanceSheet> instanceSheets = TestProgram.IgxlWorkBk.InsSheets;
                    new ConcurrentSequenceMain().WorkFlow(concurrentSheet, flowSheets, instanceSheets);
                }

                //Gen TempMon
                new TempMonMain(EpWorkbook.TestPlanWorkbook.Worksheets["TempMon_configuration"], LocalSpecs.TempMonDatas, TestProgram.NonIgxlSheetsList, FolderStructure.DirCommonSheets).WorkFlow();

                new GenNonIgxlSheetsMain().WorkFlow();

                Response.Report("Removing unnecessary if in flow ...", percentage: 85);
                foreach (KeyValuePair<string, SubFlowSheet> sheet in TestProgram.IgxlWorkBk.SubFlowSheets)
                {
                    sheet.Value.RemoveTheSameIf();
                }

                //Handle Duplicate Instance
                Response.Report("Removing Duplicate Instance ...", percentage: 90);
                foreach (KeyValuePair<string, InstanceSheet> sheet in TestProgram.IgxlWorkBk.InsSheets)
                {
                    sheet.Value.RemoveDuplicateInstance();
                }

                var duplicateInstanceChecker = new DuplicateInstanceChecker();
                duplicateInstanceChecker.WorkFlow();

                //Divide flow when row > 100000
                Response.Report("Dividing Sub Flow Sheets ...", percentage: 95);
                (Dictionary<string, SubFlowSheet> subFlowSheets, Dictionary<string, List<string>> expSheets) = DivideFlow.SplitFlowSheet(TestProgram.IgxlWorkBk.SubFlowSheets, LocalSpecs.Options.SafetyLineRowLimit);
                foreach (KeyValuePair<string, SubFlowSheet> sheet in subFlowSheets)
                {
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(sheet);
                }
                DivideFlow.InsExpFlowSheet(TestProgram.IgxlWorkBk.SubFlowSheets, expSheets);
                Dictionary<string, SubFlowSheet> subFlowDic = TestProgram.IgxlWorkBk.MainFlowSheets
                    .Where(kvp => kvp.Value is SubFlowSheet)
                    .ToDictionary(
                        kvp => kvp.Key,
                        kvp => (SubFlowSheet)kvp.Value,
                        StringComparer.OrdinalIgnoreCase // Matches your IgnoreCase setting
                    );
                DivideFlow.InsExpFlowSheet(subFlowDic, expSheets);

                ModifyBinCutFlowLimitIfNeeded();
                CheckIgxlSheetNameLength();
                CheckBinTableSheets();

                if (postActionParaData.FlagEcidSortingGen && ((TestPlanStatic.EfuseInstanceSheets != null && !TestPlanStatic.EfuseInstanceSheets.Any()) || TestPlanStatic.EfuseInstanceSheets == null))
                {
                    GenEcidSorting();
                }

                //Update TimeDomain in all PatSet sheets
                PatSetSheet patSetsAllSheet = TestProgram.IgxlWorkBk.PatSetSheets.Values.FirstOrDefault(x => x.Name.Equals("PatSets_All", StringComparison.OrdinalIgnoreCase));
                IEnumerable<PatSetSheet> patSetsSheets = TestProgram.IgxlWorkBk.PatSetSheets.Values.Where(x => !x.Name.Equals("PatSets_All", StringComparison.OrdinalIgnoreCase));
                IEnumerable<TimeSetBasicSheet> timeSetSheets = TestProgram.IgxlWorkBk.TimeSetSheets.Values;
                var updateTimeDomainMain = new UpdateTimeDomainMain(TestPlanStatic.Equipments);
                updateTimeDomainMain.UpdatePatternSetSheets(ref patSetsAllSheet, ref patSetsSheets);
                updateTimeDomainMain.UpdateTimeSetSheets(ref timeSetSheets);

                //Generate IDS and Leakage pattern summary
                Dictionary<string, PatternData> patternData = AcTSetCategoryMapSingleton.Instance().PatternList;
                if (patternData.Any())
                {
                    Response.Report("Generating IDS and Leakage Pattern Summary ...", percentage: 95);
                    new GenIdsLeakagePatSummary().WorkFlow(patternData, patInfo.SelectMany(x => x.Value).ToList(), TestProgram.IgxlWorkBk.PatSetSheets);
                }

                GenInstCommon(postActionParaData);


                SplitPatSet main = new SplitPatSet(80000);
                main.SplitPatset(TestProgram.IgxlWorkBk.PatSetSheets);


                //Gen JobList
                Response.Report("Generating JobList Sheets ...", percentage: 40);
                GenJobList();

                //Gen AAAA_TS0 into TestProgram
                AddTs0();


                Response.Report("Generating Flags into Flow_Table_Main_Init_Flags/Flow_Table_Main_Init_EnableWd ...", percentage: 95);
                var flowTableMainInitEnableWdMain = new FlowTableMainInitEnableWdMain();
                flowTableMainInitEnableWdMain.WorkFlow();
                Response.Report("Generating init items into Flow_Table_Main_Init_Flow ...", percentage: 95);
                AddInitItemsIntoInitFlow();

                if (LocalSpecs.Options.ConvertJobToEnableWord)
                {
                    ConvertJobToEnableWord(TestProgram.IgxlWorkBk.SubFlowSheets, TestProgram.IgxlWorkBk.MainFlowSheets);
                }
                //Add test Number
                Response.Report("Adding Test Number ...", percentage: 70);
                GenTestNumber(postActionParaData);

                if (TestPlanStatic.UfInstanceTable != null && TestPlanStatic.UfInstanceTable.GetErrors().Any())
                {
                    TestPlanStatic.UfInstanceTable.AddToErrorReport();
                }
            }
            catch (Exception e)
            {
                string message = "Post-action AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }

        private void GenCharSetupSheets()
        {
            foreach (CharSheet charSheet in CharSetupSheetSingleton.Instance().CharSheets)
            {
                if (charSheet.Rows.Count != 0)
                {
                    TestProgram.IgxlWorkBk.AddCharSheet(FolderStructure.DirDevChar, charSheet);
                }
            }
        }

        private void AddFlowPostIdsEqnVoltage()
        {
            if (LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                Response.Report("Adding Flow_Post_IDS_EQN_Voltage ...", percentage: 45);
                ExcelWorksheet flowPostIdsEqn = SettingStatic.BasicConfigWorkbook.Worksheets["Flow_Post_IDS_EQN_Voltage"];
                if (flowPostIdsEqn != null)
                {
                    var flowReader = new ReadFlowSheet();
                    SubFlowSheet subflow = flowReader.ReadSheet(flowPostIdsEqn);
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirMain, subflow);
                }
            }
        }

        private void GenHarvestBinTable()
        {
            if (TestPlanStatic.HarvestingTruthTableSheets != null && TestPlanStatic.HarvestingTruthTableSheets.Any())
            {
                HarvestingTruthTableSheet harvestingTruthTableSheet = TestPlanStatic.HarvestingTruthTableSheets.First();
                Response.Report("Generating Harvest BinTable ...", percentage: 75);
                BinTableSheet binTable = TestProgram.IgxlWorkBk.GetBinTblSheet(FolderStructure.DirNonBinCut, "Bin_Table_Non_BinCut");
                List<BinTableRow> binTableRows = new ScanHarvestMain(harvestingTruthTableSheet).GenBinTableRows();
                binTable.AddRows(binTableRows.FindAll(x => !binTable.Rows.Exists(y => y.Name.Equals(x.Name))));
            }
        }

        private void GenEcidSorting()
        {
            #region EFUSE ECID Sorting
            SubFlowSheet mainInitEnableWdSheet = TestProgram.IgxlWorkBk.GetFlowSheet(IgxlWorkBook.FlowTableMainInitEnableWd, FolderStructure.DirMain);
            SubFlowSheet ecidSubFlowSheet = TestProgram.IgxlWorkBk.GetFlowSheet(EFuseConst.EcidFlowSheet, FolderStructure.DirEfuse);
            SubFlowSheet ecidDeidSubFlowSheet = TestProgram.IgxlWorkBk.GetFlowSheet(EFuseConst.EcidDeidFlowSheet, FolderStructure.DirEfuse);
            InstanceSheet ecidInstanceSheet = TestProgram.IgxlWorkBk.GetInstanceSheet(EFuseConst.TestInstSheet);
            if (!ecidSubFlowSheet.Rows.Count.Equals(0))
            {
                EcidProgramGenerateMain.GenSyntaxCheckFlow(ecidSubFlowSheet);
                EcidProgramGenerateMain.GenEcidSortingFlow(ecidSubFlowSheet);
            }
            if (!ecidDeidSubFlowSheet.Rows.Count.Equals(0))
            {
                EcidProgramGenerateMain.GenSyntaxCheckFlow(ecidDeidSubFlowSheet);
                EcidProgramGenerateMain.GenEcidSortingFlow(ecidDeidSubFlowSheet);
            }
            if (ecidInstanceSheet != null)
            {
                EcidProgramGenerateMain.GenEcidSortingInst(ecidInstanceSheet);
            }
            EcidProgramGenerateMain.GenEnableToMainInitEnableWd(mainInitEnableWdSheet, "ECIDSort_Enable", OpCode.DisableFlowWd, EFuseConst.Efuse);
            EcidProgramGenerateMain.GenEnableToMainInitEnableWd(mainInitEnableWdSheet, "Enable_ECID_Sorting_2CMode", OpCode.DisableFlowWd, EFuseConst.Efuse);
            #endregion
        }

        private void ModifyBinCutFlowLimitIfNeeded()
        {
            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.BinCut).Down && LocalSpecs.IsModuleIncluded(BlockStatus.BinCut))
            {
                if (BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && LocalSpecs.IsModuleIncluded(BlockStatus.HardIp))
                {
                    Response.Report("Modifying limit in flow for BinCut ...", percentage: 95);
                    ModifyBinCutFlowLimit();
                }
            }
        }

        private void CheckBinTableSheets()
        {
            try
            {
                foreach (KeyValuePair<string, BinTableSheet> sheet in TestProgram.IgxlWorkBk.BinTblSheets)
                {
                    sheet.Value.Check();
                }
            }
            catch (Exception e)
            {
                Response.Report(e.StackTrace, EnumMessageLevel.Error, 0);
            }
        }

        private void GenInstCommon(PostActionParaData postActionParaData)
        {
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) && Directory.Exists(LocalSpecs.CsLibraryFolder))
            {
                var removeHeader = new RemoveHeader();
                removeHeader.WorkFlow();
                var instCommon = new InstCommonMain();
                instCommon.WorkFlowWithoutHeader(postActionParaData.TestInstCommon);
            }
            else
            {
                var instCommon = new InstCommonMain();
                instCommon.WorkFlowWithHeader(postActionParaData.TestInstCommon);
            }
        }

        private void GenJobList()
        {
            var jobListGen = new JobListMain();
            if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any())
            {
                foreach (TestProgramRow testProgramRow in TestPlanStatic.TestProgramDefSheet.Rows)
                {
                    jobListGen.WorkFlow(testProgramRow);
                }
            }
            else
            {
                jobListGen.WorkFlow();
            }
        }

        public static PostActionParaData GetPostActionParas()
        {
            Response.Report("Running Post-ModuleMain~", EnumMessageLevel.CheckPoint);
            var postActionParaData = new PostActionParaData();
            postActionParaData.TestInstCommon = "TestInst_Common";

            if (!BlockStatus.GetAutomationBlockStatus(BlockStatus.Efuse).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Efuse))
            {
                string eFuseTemplate = Path.Combine(LocalSpecs.SettingFolder,
                                       "Settings", "eFuse", $"Efuse_Flow_TestInst_{LocalSpecs.CurrentProject}.xlsx");
                var inputExcel = new ExcelPackage(new FileInfo(eFuseTemplate));
                EpWorkbook.EfuseTemplateWorkbook = inputExcel.Workbook;
            }
            else
            {
                postActionParaData.InitMainFlowFlag = false;  //if efuse is enabled , disable this one.
            }

            string tn = Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic", $"TN_Assignment_{LocalSpecs.CurrentProject}.xlsx");
            if (!string.IsNullOrEmpty(LocalSpecs.SettingFiles.TnAssignment) && File.Exists(LocalSpecs.SettingFiles.TnAssignment))
            {
                postActionParaData.TestNumberXlsx = LocalSpecs.SettingFiles.TnAssignment;
                Response.Report($"{LocalSpecs.SettingFiles.TnAssignment} was specified!!");
            }
            else if (File.Exists(tn))
            {
                postActionParaData.TestNumberXlsx = tn;
                Response.Report($"TN_Assignment_{LocalSpecs.CurrentProject}.xlsx was specified!!");
            }
            else
            {
                postActionParaData.TestNumberXlsx = "N/A";
                Response.Report($"TN_Assignment_{LocalSpecs.CurrentProject}.xlsx Not Found??", EnumMessageLevel.Warning);
            }

            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Efuse).Down && LocalSpecs.IsModuleIncluded(BlockStatus.Efuse))
            {
                postActionParaData.FlagEcidSortingGen = true;
            }
            return postActionParaData;
        }

        private void AddTs0()
        {
            foreach (KeyValuePair<string, TimeSetBasicSheet> keyValuePair in TestProgram.IgxlWorkBk.TimeSetSheets)
            {
                var addTsList = new List<TSet>();
                var addPinList = new List<string>();
                foreach (TSet tSet in keyValuePair.Value.Rows)
                {
                    if (tSet.Name == "")
                    {
                        break;
                    }

                    List<TimingRow> timingRows = tSet.TimingRows;
                    var newTs = new TSet
                    {
                        Name = "AAAA_TS0",
                        CyclePeriod = tSet.CyclePeriod,
                    };
                    foreach (TimingRow timingRow in timingRows)
                    {
                        if (timingRow.PinGrpName == null)
                        {
                            continue;
                        }

                        if (timingRow.DataSrc.Equals("PA", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (addPinList.Exists(x => x.Equals(timingRow.PinGrpName, StringComparison.OrdinalIgnoreCase)))
                        {
                            continue;
                        }

                        TimingRow copy = timingRow.Copy();
                        addPinList.Add(copy.PinGrpName);
                        bool is2X = copy.PinGrpSetup.ContainsIgnoreCase("_2X");
                        copy.DataFmt = is2X ? "STAY-2X" : "STAY";
                        copy.DriveOff = "0";
                        copy.DriveOn = "";
                        copy.DriveData = "";
                        copy.DriveReturn = "";
                        newTs.AddTimingRow(copy);
                    }
                    if (newTs.TimingRows.Any())
                    {
                        addTsList.Add(newTs);
                    }
                }
                int blankIdx = keyValuePair.Value.Rows.FindIndex(x => string.IsNullOrEmpty(x.Name));
                foreach (TSet tSet in addTsList)
                {
                    if (blankIdx != -1)
                    {
                        keyValuePair.Value.InsertTimeSet(blankIdx++, tSet);
                    }
                    else
                    {
                        keyValuePair.Value.AddRow(tSet);
                    }
                }
                if (keyValuePair.Key.Contains(".txt"))
                {
                    keyValuePair.Value.Write(keyValuePair.Key);
                }
                else
                {
                    keyValuePair.Value.Write(keyValuePair.Key + ".txt");
                }
            }
        }

        private void AddInitItemsIntoInitFlow()
        {
            SubFlowSheet mainInitFlowSheet = TestProgram.IgxlWorkBk.GetFlowSheet(IgxlWorkBook.FlowTableMainInitFlows, FolderStructure.DirMain);
            if (mainInitFlowSheet == null)
            {
                return;
            }

            MoveHardipTtrEnableToInitFlow(mainInitFlowSheet);
            GenerateEnableWdGating(mainInitFlowSheet);
            TestProgram.IgxlWorkBk.SetFlowSheet(IgxlWorkBook.FlowTableMainInitFlows, mainInitFlowSheet);
        }

        private void MoveHardipTtrEnableToInitFlow(SubFlowSheet initFlow)
        {
            SubFlowSheet mainInitEnableWdSheet = TestProgram.IgxlWorkBk.GetFlowSheet(IgxlWorkBook.FlowTableMainInitEnableWd, FolderStructure.DirMain);
            InstanceSheet instSheet = TestProgram.IgxlWorkBk.InsSheets.Values.ToList().Find(p => p.Name.Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));
            if (instSheet == null)
            {
                return;
            }
            if (mainInitEnableWdSheet != null)
            {
                int instIdx = mainInitEnableWdSheet.Rows.FindIndex(x => x.Parameter.Equals("HIP_TTR_Enable_Control", StringComparison.OrdinalIgnoreCase));
                if (instIdx > 0)
                {
                    mainInitEnableWdSheet.Rows.RemoveAt(instIdx);
                    TestProgram.IgxlWorkBk.SetFlowSheet(IgxlWorkBook.FlowTableMainInitEnableWd, mainInitEnableWdSheet);
                }
            }
            if (!instSheet.Rows.Exists(x => x.TestName.Equals("HIP_TTR_Enable_Control", StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }
            if (!initFlow.Rows.Exists(x => x.Parameter.Equals("HIP_TTR_Enable_Control", StringComparison.OrdinalIgnoreCase)))
            {
                var flowRows = new FlowRows { new FlowRow { Opcode = OpCode.Test, Parameter = "HIP_TTR_Enable_Control" } };
                initFlow.InsertBeforeReturnRow(flowRows);
            }
        }

        private void GenerateEnableWdGating(SubFlowSheet initFlow)
        {
            if (initFlow != null && SettingStatic.BasicConfigWorkbook.Worksheets["EnableWdGatingTable"] != null)
            {
                var rows = new List<FlowRow>
                {
                    new FlowRow { Opcode = OpCode.Test, Parameter = "EnableWdGating", FailAction = "F_ecid_other" },
                    new FlowRow { Opcode = OpCode.BinTable, Parameter = "Bin_EFUSE_ecid_other" }
                };
                initFlow.InsertBeforeReturnRow(rows);
                bool isCsharp = TestProgram.VbtFunctionLib.GetFunctionByName("EnableWordGatingFunc", "scan", true) != null;
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName("EnableWdGating", "scan");
                if (function != null && !isCsharp)
                {
                    InstanceSheet instSheet = TestProgram.IgxlWorkBk.InsSheets.Values.ToList()
                        .Find(p => p.Name.Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));

                    var row = new InstanceRow
                    {
                        TestName = "EnableWdGating",
                        VbtType = "VBT",
                        VbtName = "EnableWdGating"
                    };
                    function.SetParamValue("realPhase", "3");
                    row.ArgList = function.Parameters;
                    row.Args = function.ArgList;
                    instSheet?.AddRow(row);
                }
            }
        }

        private void ModifyBinCutFlowLimit()
        {
            Response.Report("Modify BinCut Flow Limit for ELB ...", percentage: 5);
            InstanceSheet testInstSheet = TestProgram.IgxlWorkBk.GetInstanceSheet("TestInst_VddBinning");
            var flowSheets = TestProgram.IgxlWorkBk.SubFlowSheets.Select(x => x.Value).ToList();
            var instSheets = TestProgram.IgxlWorkBk.InsSheets.Select(x => x.Value).ToList();
            var rtosList = new List<string>();
            if (EpWorkbook.BinCutWorkbook == null)
            {
                return;
            }

            ExcelWorksheet binCutWorksheet = EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BcFlow] ?? EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BcAte];
            if (binCutWorksheet != null)
            {
                var binCutSheetReader = new BinCutFlowSheetReader();
                BinCutFlowTables binCutFlowTables = binCutSheetReader.ReadSheet(binCutWorksheet);
                //ErrorReportManager.AddErrors(binCutFlowTables.SelectMany(x => x.Errors).ToList());
                foreach (BinCutFlowTable sheet in binCutFlowTables)
                {
                    foreach (BinCutFlowSheetRow row in sheet.Rows)
                    {
                        if (row.SpiRtos.Contains("_"))
                        {
                            rtosList.Add(row.SpiRtos);
                        }
                    }
                }
            }
            if (testInstSheet != null && flowSheets.Any())
            {
                new BinCutElbMain().WorkFlow(instSheets, flowSheets, rtosList.Distinct().ToList());
            }
        }

        private void OptimizePatSetAll()
        {
            IEnumerable<PatSet> allPatSetSheets = TestProgram.IgxlWorkBk.PatSetSheets.Where(x => !x.Value.Name.Equals("PatSets_All")).SelectMany(x => x.Value.Rows);
            var allPatSetRows = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (PatSet sheet in allPatSetSheets)
            {
                if (allPatSetRows.ContainsKey(sheet.PatSetName))
                {
                    Response.Report($"[Error] Found duplicate PatSetName : {sheet.PatSetName}", EnumMessageLevel.Error);
                }
                else
                {
                    allPatSetRows.Add(sheet.PatSetName, sheet.PatSetRows.Select(y => y.File).ToHashSet(StringComparer.OrdinalIgnoreCase));
                }
            }

            var actualUsedPattern = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var patSetAll = TestProgram.IgxlWorkBk.PatSetSheets.Where(x => x.Value.Name.Equals("PatSets_All")).SelectMany(x => x.Value.Rows).Select(x => x.PatSetName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            PatSetSheet patSetAllRow = TestProgram.IgxlWorkBk.GetPatSetsSheet(IgxlWorkBook.PatSetsAll, FolderStructure.DirPatSetsAll);
            PatSetSubSheet patSubSheet = TestProgram.IgxlWorkBk.GetPatSetSubSheet(IgxlWorkBook.PatternSubroutine, FolderStructure.DirPatSetsAll);
            var patSetDashBoard = new PatSetSheet("PatSets_DashBoard");
            var removePatSubList = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            patSetDashBoard.AddRows(patSetAllRow.Rows);
            TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirPatSetsAll, patSetDashBoard);
            foreach (string pat in VbtFunctionLibShared.UsedPatternList)
            {
                if (actualUsedPattern.Contains(pat))
                {
                    continue;
                }

                if (allPatSetRows.ContainsKey(pat.ToUpper()))
                {
                    actualUsedPattern.AddRange(allPatSetRows[pat.ToUpper()]);
                }
                else
                {
                    actualUsedPattern.Add(pat.ToUpper());
                }
            }
            var nonUsedPat = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string pat in patSetAll)
            {
                if (pat.ContainsIgnoreCase("MLPSUB"))
                {
                    continue;
                }

                if (!actualUsedPattern.Contains(pat))
                {
                    nonUsedPat.Add(pat.ToUpper());
                }
            }
            Dictionary<string, PatSet> allExistPatSetDic = patSetAllRow.PatSetRowDic;
            foreach (string pat in nonUsedPat)
            {
                PatSet existItem = allExistPatSetDic.TryGetValue(pat, out PatSet value) ? value : null;
                if (existItem != null)
                {
                    patSetAllRow.Rows.Remove(existItem);
                    allExistPatSetDic.Remove(pat);
                }
            }
            var realPatternPath = patSetAllRow.Rows.SelectMany(x => x.PatSetRows).Select(x => x.File.Split(':')[0]).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (PatSetSubRow patSetSubRow in patSubSheet.Rows)
            {
                if (!string.IsNullOrEmpty(patSetSubRow.ColumnA) && patSetSubRow.ColumnA.Equals("SubrNoVm"))
                {
                    continue;
                }

                string patternName = patSetSubRow.PatternFileName.Split(':')[0].ToUpper();
                if (!patternName.Contains("MLPSUB") && !realPatternPath.Contains(patternName))
                {
                    removePatSubList.Add(patSetSubRow.PatternFileName);
                }
            }
            foreach (string pat in removePatSubList)
            {
                PatSetSubRow existItem = patSubSheet.GetExist(pat.ToUpper());
                patSubSheet.Rows.Remove(existItem);
            }
        }

        private void CheckIgxlSheetNameLength()
        {
            var ignoreTimeList = TestProgram.IgxlWorkBk.TimeSetSheets.Where(x => Path.GetFileName(x.Key).ToString().Length >= 32).ToList();
            if (ignoreTimeList.Count > 0)
            {
                Response.Report("[Error] The Timing set list below will not be included in test program!", EnumMessageLevel.Error);
            }

            foreach (KeyValuePair<string, TimeSetBasicSheet> item in ignoreTimeList)
            {
                string msg = item.Key + ": The length of Timing set name exceeds the maximum 31 chars ";
                Response.Report(msg, EnumMessageLevel.Error);
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_07, "", 0, 0, item.Key + ": The length of Timing set name exceeds the maximum 31 chars ");
                TestProgram.IgxlWorkBk.AllIgxlSheets.Remove(item.Key);
                TestProgram.IgxlWorkBk.TimeSetSheets.Remove(item.Key);
                TestProgram.IgxlWorkBk.RestrictedIgxlSheets.Add(item.Key, item.Value);
            }

            var ignoreFlowList = TestProgram.IgxlWorkBk.SubFlowSheets.Where(x => Path.GetFileName(x.Key).ToString().Length >= 32).ToList();
            if (ignoreFlowList.Count > 0)
            {
                Response.Report("[Error] The Flow sheets list below will not be included in test program!", EnumMessageLevel.Error);
            }

            foreach (KeyValuePair<string, SubFlowSheet> item in ignoreFlowList)
            {
                string msg = item.Key + ": The length of Flow sheet name exceeds the maximum 31 chars ";
                Response.Report(msg, EnumMessageLevel.Error);
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_08, "", 0, 0, item.Key + ": The length of Flow sheet name exceeds the maximum 31 chars ");
                TestProgram.IgxlWorkBk.AllIgxlSheets.Remove(item.Key);
                TestProgram.IgxlWorkBk.SubFlowSheets.Remove(item.Key);
                TestProgram.IgxlWorkBk.RestrictedIgxlSheets.Add(item.Key, item.Value);
            }

            var ignoreInstList = TestProgram.IgxlWorkBk.InsSheets.Where(x => Path.GetFileName(x.Key).ToString().Length >= 32).ToList();
            if (ignoreInstList.Count > 0)
            {
                Response.Report("[Error] The Inst sheets list below will not be included in test program!", EnumMessageLevel.Error);
            }

            foreach (KeyValuePair<string, InstanceSheet> item in ignoreInstList)
            {
                string msg = item.Key + ": The length of Inst sheet name exceeds the maximum 31 chars ";
                Response.Report(msg, EnumMessageLevel.Error);
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_09, "", 0, 0, item.Key + ": The length of Inst sheet name exceeds the maximum 31 chars ");
                TestProgram.IgxlWorkBk.AllIgxlSheets.Remove(item.Key);
                TestProgram.IgxlWorkBk.InsSheets.Remove(item.Key);
                TestProgram.IgxlWorkBk.RestrictedIgxlSheets.Add(item.Key, item.Value);
            }
        }

        private void GenTestNumber(PostActionParaData postActionParaData)
        {
            var tnPara = new TestNumberParaData();
            if (postActionParaData.TestNumberXlsx == "N/A")
            {
                Response.Report("Test Number mapping sheet didn't specify, ignore and not insert the test number#.", EnumMessageLevel.Warning, 60);
            }
            else
            {
                tnPara.TnXlsx = postActionParaData.TestNumberXlsx;
                var testNumber = new TestNumberMain(LocalSpecs.SettingFiles.TnAssignment);
                List<string> nonTnSheets = testNumber.WorkFlow(tnPara);
                if (nonTnSheets.Count > 0)
                {
                    Response.Report(string.Format("SubFlow Sheets didn't have specify mapping for test number:\n" + string.Join("\n", nonTnSheets)), EnumMessageLevel.Warning, 60);
                }
            }
        }

        #region Print Flow

        public void PrintFlow()
        {
            TestProgram.IgxlWorkBk.PrintAllSheets();
            TestProgram.SubProgIgxlWorkBk.PrintAllSheets();
            foreach (string customPath in LocalSpecs.CustomPath)
            {
                CopyFromExtraSheets(customPath);
            }
        }

        #endregion
        private void GenItemsFromExtra()
        {
            var insSheets = TestProgram.IgxlWorkBk.InsSheets.Select(x => x.Value.Name).ToList();
            var flowSheets = TestProgram.IgxlWorkBk.SubFlowSheets.Select(x => x.Value.Name).ToList();
            foreach (string customPath in LocalSpecs.CustomPath)
            {
                AddFlowSheetByExtraSheet(customPath, flowSheets);
                AddInstanceSheetByExtraSheet(customPath, insSheets);
            }
        }
        private void AddInstanceSheetByExtraSheet(string extraFolder, List<string> insSheets)
        {
            var extraInstances = new List<string>();
            if (Directory.Exists(extraFolder))
            {
                extraInstances = IgxlSheetReaderHelpers.GetSheetsByType(extraFolder, EnumSheetType.DTTestInstancesSheet);
                insSheets.AddRange(extraInstances.Select(Path.GetFileNameWithoutExtension));

            }
        }

        private void AddFlowSheetByExtraSheet(string extraFolder, List<string> flowSheets)
        {
            var extraSubflows = new List<string>();
            if (Directory.Exists(extraFolder))
            {
                extraSubflows = IgxlSheetReaderHelpers.GetSheetsByType(extraFolder, EnumSheetType.DTFlowtableSheet);
                flowSheets.AddRange(extraSubflows.Select(Path.GetFileNameWithoutExtension));
            }
        }

        internal void CopyFromExtraSheets(string extraFolder)
        {
            if (Directory.Exists(extraFolder))
            {
                var dir = new DirectoryInfo(extraFolder);
                FileInfo[] files = dir.GetFiles("*.txt");
                foreach (FileInfo file in files)
                {
                    if (Regex.IsMatch(file.Name, "^Channel*", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    bool flag = false;
                    foreach (KeyValuePair<string, IIgxlSheet> igxlSheetPair in TestProgram.IgxlWorkBk.AllIgxlSheets)
                    {
                        string fileName = Path.GetFileName(igxlSheetPair.Key + ".txt");
                        if (fileName.Equals(file.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            flag = true;
                            if (IgxlSheetReaderHelpers.IsSheetType(EnumSheetType.DTFlowtableSheet, file.FullName))
                            {
                                KeepOriginalSubFlow(igxlSheetPair);
                            }

                            if (IgxlSheetReaderHelpers.IsSheetType(EnumSheetType.DTTestInstancesSheet, file.FullName))
                            {
                                ReplaceInstanceByExtra(igxlSheetPair, file.FullName);
                                break;
                            }
                            if (IgxlSheetReaderHelpers.IsSheetType(EnumSheetType.DTPatternSetSheet, file.FullName))
                            {
                                if (Regex.IsMatch(file.Name, "^Patsets_all*", RegexOptions.IgnoreCase))
                                {
                                    continue;
                                }

                                ReplacePatSetByExtra(igxlSheetPair, file.FullName);
                                break;
                            }
                            File.Copy(file.FullName, igxlSheetPair.Key + ".txt", true);
                            string outString = $"ExtraSheet:{file.Name} Conflict with generate output... Please Check";
                            Response.Report(outString, EnumMessageLevel.Warning, 0);
                            break;
                        }
                    }
                    if (!flag)
                    {
                        if (!_existedExtraSheet.Contains(file.Name))
                        {
                            _existedExtraSheet.Add(file.Name);
                            File.Copy(file.FullName, Path.Combine(FolderStructure.DirCommonSheets, file.Name), true);
                            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, file.Name.Replace(".txt", ""));
                        }
                        else
                        {
                            string outString = $"ExtraSheet:{file.Name} already added into output... Skip";
                            Response.Report(outString, EnumMessageLevel.Warning, 0);
                        }
                    }
                }
            }
        }

        private void KeepOriginalSubFlow(KeyValuePair<string, IIgxlSheet> igxlSheetPair)
        {
            string filePath = Path.GetDirectoryName(igxlSheetPair.Key);
            string originalFileName = Path.GetFileName(igxlSheetPair.Key);
            string newFileName = originalFileName + "_autogen";
            if (filePath != null)
            {
                File.Move(Path.Combine(filePath, originalFileName) + ".txt",
                    Path.Combine(filePath, newFileName) + ".txt");
                SubFlowSheet newSubflowSheet =
                    new ReadFlowSheet().GetSheet(Path.Combine(filePath, newFileName) + ".txt");
                newSubflowSheet.Name = newFileName;
                TestProgram.IgxlWorkBk.AddSubFlowSheet(filePath, newSubflowSheet);
            }
        }

        private void ReplacePatSetByExtra(KeyValuePair<string, IIgxlSheet> igxlSheetPair, string extraFile)
        {
            string oriSheet = Path.GetFileName(igxlSheetPair.Key);
            PatSetSheet extraInstanceSheet = new ReadPatSetSheet().GetSheet(extraFile);
            string filePath = Path.GetDirectoryName(igxlSheetPair.Key);
            PatSetSheet patSetSheet = TestProgram.IgxlWorkBk.PatSetSheets.ToList().Find(x => x.Value.Name.Equals(oriSheet, StringComparison.OrdinalIgnoreCase)).Value;
            foreach (PatSet extraPatSet in extraInstanceSheet.Rows)
            {
                PatSet existTargetPatSet = patSetSheet.Rows.Find(x => x.PatSetName.Equals(extraPatSet.PatSetName));

                if (existTargetPatSet != null)
                {
                    if (!PatSetCompare(extraPatSet, existTargetPatSet))
                    {
                        int index = patSetSheet.Rows.FindIndex(x => x.PatSetName.Equals(extraPatSet.PatSetName));
                        patSetSheet.Rows[index].IsBackup = true;
                        patSetSheet.Rows[index].PatSetRows.ForEach(x => x.IsBackup = true);
                        patSetSheet.Rows[index].PatSetRows.ForEach(x => x.ColumnA += "Autogen");
                        extraPatSet.PatSetRows.ForEach(x => x.ColumnA += "Custom");
                        patSetSheet.Rows.Insert(index, extraPatSet);
                    }

                }
                else
                {
                    extraPatSet.PatSetRows.ForEach(x => x.ColumnA += "Custom");
                    patSetSheet.Rows.Add(extraPatSet);
                }
            }

            if (filePath != null)
            {
                patSetSheet.Write(Path.Combine(filePath, patSetSheet.Name) + ".txt");
                TestProgram.IgxlWorkBk.AddPatSetSheet(filePath, patSetSheet);
            }
        }

        private void ReplaceInstanceByExtra(KeyValuePair<string, IIgxlSheet> igxlSheetPair, string extraFile)
        {
            string oriSheet = Path.GetFileName(igxlSheetPair.Key);
            InstanceSheet extraInstanceSheet = new ReadInstanceSheet().GetSheet(extraFile);
            string filePath = Path.GetDirectoryName(igxlSheetPair.Key);
            InstanceSheet newInstanceSheet = TestProgram.IgxlWorkBk.InsSheets.ToList().Find(x => x.Value.Name.Equals(oriSheet, StringComparison.OrdinalIgnoreCase)).Value;

            foreach (InstanceRow extraInstance in extraInstanceSheet.Rows)
            {
                InstanceRow existTargetInstance = newInstanceSheet.Rows.Find(x => x.TestName.Equals(extraInstance.TestName));

                if (existTargetInstance != null)
                {
                    int argCount = existTargetInstance.ArgList.Split(',').Length;
                    existTargetInstance.Args.RemoveRange(argCount, existTargetInstance.Args.Count - argCount);
                    int exTraArgCount = extraInstance.ArgList.Split(',').Length;
                    extraInstance.Args.RemoveRange(exTraArgCount, extraInstance.Args.Count - exTraArgCount);
                    if (!InstanceCompare(extraInstance, existTargetInstance))
                    {
                        int index = newInstanceSheet.Rows.FindIndex(x => x.TestName.Equals(extraInstance.TestName));
                        newInstanceSheet.Rows[index].IsBackup = true;
                        newInstanceSheet.Rows[index].ColumnA += " ,Autogen";
                        newInstanceSheet.Rows.Insert(index, extraInstance);
                    }
                }
                else
                {
                    extraInstance.ColumnA += " ,Custom";
                    newInstanceSheet.Rows.Add(extraInstance);
                }
            }

            if (filePath != null)
            {
                newInstanceSheet.Write(Path.Combine(filePath, newInstanceSheet.Name) + ".txt");
                TestProgram.IgxlWorkBk.AddInsSheet(filePath, newInstanceSheet);
            }
        }

        internal bool PatSetCompare(PatSet row1, PatSet row2)
        {
            bool result = true;

            row1.GetType().GetProperties().ToList().ForEach(p =>
            {
                row2.GetType().GetProperties().ToList().ForEach(p2 =>
                {
                    if (p.Name == p2.Name && !p.Name.Equals("RowNum") && !p.Name.Equals("Domain"))
                    {
                        if (p.Name.Equals("PatSetRows"))
                        {
                            if (row1.PatSetRows.Count != row2.PatSetRows.Count)
                            {
                                result = false;
                            }
                            else
                            {
                                int rowCount = row1.PatSetRows.Count;
                                for (int i = 0; i < rowCount; i++)
                                {
                                    if (!row1.PatSetRows[i].CompareRow(row2.PatSetRows[i]))
                                    {
                                        result = false;
                                    }
                                }
                            }
                        }
                        else
                        {
                            object compare1 = p.GetValue(row1) ?? "";
                            object compare2 = p2.GetValue(row2) ?? "";
                            if (!compare1.Equals(compare2))
                            {
                                result = false;
                            }
                        }
                    }
                });
            });
            return result;
        }

        internal bool InstanceCompare(InstanceRow row1, InstanceRow row2)
        {
            bool result = true;

            row1.GetType().GetProperties().ToList().ForEach(p =>
            {
                row2.GetType().GetProperties().ToList().ForEach(p2 =>
                {
                    if (p.Name == p2.Name && !p.Name.Equals("RowNum"))
                    {
                        if (p.Name.Equals("Args"))
                        {
                            if (!string.Join("+", row1.Args).Equals(string.Join("+", row2.Args)))
                            {
                                result = false;
                            }
                        }
                        else if (p.Name.Equals("InitList"))
                        {
                            if (!string.Join("+", row1.InitList).Equals(string.Join("+", row2.InitList)))
                            {
                                result = false;
                            }
                        }
                        else if (p.Name.Equals("PayloadList"))
                        {
                            if (!string.Join("+", row1.PayloadList).Equals(string.Join("+", row2.PayloadList)))
                            {
                                result = false;
                            }
                        }
                        else if (p.Name.Equals("FinalJobs"))
                        {
                            if (!string.Join("+", row1.FinalJobs).Equals(string.Join("+", row2.FinalJobs)))
                            {
                                result = false;
                            }
                        }
                        else if (!p.GetValue(row1).Equals(p2.GetValue(row2)))
                        {
                            result = false;
                        }
                    }
                });
            });
            return result;
        }

        internal void ConvertJobToEnableWord(Dictionary<string, SubFlowSheet> subFlowSheets, Dictionary<string, MainFlow> mainFlowSheets)
        {
            foreach (KeyValuePair<string, SubFlowSheet> subflow in subFlowSheets)
            {
                if (subflow.Value.Name.Equals("Flow_Table_Main_Init_EnableWd", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                ConvertJobRowsToEnableWord(subflow.Value.Rows);
            }

            foreach (KeyValuePair<string, MainFlow> subflow in mainFlowSheets)
            {
                if (subflow.Value.Name.Equals(MainFlowMain.SubProgramMainFlow, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                ConvertJobRowsToEnableWord(subflow.Value.Rows);
            }
        }

        private void ConvertJobRowsToEnableWord(IEnumerable<FlowRow> rows)
        {
            IEnumerable<FlowRow> items = rows.Where(p => !p.Opcode.Equals("use-limit", StringComparison.OrdinalIgnoreCase) && !p.Opcode.Equals("if", StringComparison.OrdinalIgnoreCase) && !p.Opcode.Equals("endif", StringComparison.OrdinalIgnoreCase)
                                                         && !string.IsNullOrEmpty(p.Job));
            foreach (FlowRow item in items)
            {
                List<string> jobs = item.Job.Split(',').ToList();
                string jobEnables = jobs.Exists(x => x.Contains("!")) ? string.Join("&&", jobs) : string.Join("||", jobs);
                if (string.IsNullOrEmpty(item.Enable))
                {
                    item.Enable = jobEnables;
                }
                else
                {
                    item.Enable = $"({item.Enable})&&({jobEnables})";
                }

                item.Job = "";
            }
        }
    }
}
