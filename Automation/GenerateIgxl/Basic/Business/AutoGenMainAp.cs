using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenAc;
using Automation.GenerateIgxl.Basic.Business.GenAc.AcGenerator.Business;
using Automation.GenerateIgxl.Basic.Business.GenConti.Base;
using Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy;
using Automation.GenerateIgxl.Basic.Business.GenDc;
using Automation.GenerateIgxl.Basic.Business.GenGlobalDc.Business;
using Automation.GenerateIgxl.Basic.Business.GenGlobalSpec;
using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;
using Automation.GenerateIgxl.Basic.Business.GenLevel.Business;
using Automation.GenerateIgxl.Basic.Business.GenMappingTable;
using Automation.GenerateIgxl.Basic.Business.GenNonIgxlSheet;
using Automation.GenerateIgxl.Basic.Business.GenNwire.Base;
using Automation.GenerateIgxl.Basic.Business.GenNwire.Business;
using Automation.GenerateIgxl.Basic.Business.GenNwire.Business.CopyXml;
using Automation.GenerateIgxl.Basic.Business.GenPatSet.Business;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Override.Applier;
using Automation.GenerateIgxl.PostAction.Relay;
using Automation.PreCheck.AllParaData;
using Automation.Reader;
using Automation.Reader.ConfigFile.RtosTable;
using Automation.Singleton;
using Automation.Static;
using Automation.Static.Result;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.DataStruct;

namespace Automation.GenerateIgxl.Basic.Business
{
    public class AutogenMainAp : BasicMain
    {
        protected int ProcessStatus;

        #region Src
        protected List<PatternData> PatList;
        #endregion

        #region Mid
        protected BasicInitial Initial;
        protected BasicMidResult Mid = new BasicMidResult();
        #endregion

        #region Tar
        protected ContiResult ContiResult;
        protected NwireResult NwireResult;
        #endregion

        protected MultiTestSettingSheetsSingleton MultiTestSettingSheetsSingleton { set; get; }
        private List<string> _specialTsValueFiles;

        public override void WorkFlow(ParaData paraData)
        {
            ProcessStatus = 10;
            try
            {
                BasicInitial();

                GenNonIgxlSheet();

                GenLevel();

                GenGlobalSpec();

                GenDcSpec(MultiTestSettingSheetsSingleton.DcCategoryInfos.IsSplitDcSpecs(LocalSpecs.Options.IsSplitDcSpecs));

                GenPatternSet();

                GenTimeSet();

                GenAcSpec();

                GenUfInstanceSheet();

                GenContinuity();

                GenNwire();

                GenMappingTable();

                AddIgxlSheet();

                GenRelay();

                AddNonIgxlSheet();

                Response.Report("Basic Completed!", percentage: ProcessStatus = 100);
            }
            catch (Exception e)
            {
                string message = "Basic AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }

        protected virtual void BasicInitial()
        {
            Response.Report("Initializing Basic ...", percentage: ProcessStatus = 10);
            Initial = new BasicInitial();
            if (BasicResult.PatternList)
            {
                Response.Report("Reading Patterns/Timing Set from PatList File ...", percentage: ProcessStatus = 12);
                PatList = AcTSetCategoryMapSingleton.Instance().PatternList.Select(x => x.Value).ToList();
            }

            CopyNwireFile();
            try
            {
                MultiTestSettingSheetsSingleton = MultiTestSettingSheetsSingleton.Instance();

                PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
                var tmpIoInfoSeq = TestPlanStatic.PowerInfoSheet?.Rows.Where(pin => pin.IsInitPin).ToList();
                if (tmpIoInfoSeq?.Count > 0)
                {
                    LocalSpecs.ExistIoSeqHiLo = true;
                    foreach (PowerInfoRow ioInfoSeq in tmpIoInfoSeq)
                    {
                        if (!pinMap.IsPinExist(ioInfoSeq.PinName))
                        {
                            pinMap.AddPin(new Pin(ioInfoSeq.PinName, CommonConst.IoPin));
                        }

                        if (ioInfoSeq.IsInitPin)
                        {
                            CreateTargetPinGroup(pinMap, ioInfoSeq.PinName, CommonConst.InitPins);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Response.Report("Meet an Error in New Basic: " + e.StackTrace, EnumMessageLevel.Error, 100);
            }
        }

        protected virtual void GenAcSpec()
        {
            if (BasicResult.Ac)
            {
                //Initial AC Specs
                Response.Report("Generate AC Specs sheet ...", percentage: ProcessStatus = 45);
                if (Mid.AcSpecSheet == null)
                {
                    Mid.AcInputSheet = Initial.InitalAcSymbols();
                    var acGenerator = new AcGenerator(Mid.AcInputSheet);
                    Mid.AcSpecSheet = acGenerator.GenerateFlow(Mid.TimeSetSheets, PatList);
                }
            }
        }

        private void CreateTargetPinGroup(PinMapSheet pinMapSheet, string pinName, string pinGroup)
        {
            if (!pinMapSheet.IsGroupExist(pinGroup))
            {
                var initGroup = new PinGroup(pinGroup, "I/O");
                initGroup.AddPin(pinName);
                pinMapSheet.AddGroup(initGroup);
            }
            else
            {
                pinMapSheet.GetGroup(pinGroup).AddPin(pinName);
            }
        }

        private void GenDcSpec(bool isSplitDcSpecs)
        {
            Response.Report("Generating DcSpec ...", percentage: ProcessStatus = 50);
            var dcGenerator = new DcSpecGenerator(MultiTestSettingSheetsSingleton, TestPlanStatic.IoInfoSheet, TestPlanStatic.IoInfoConcurrentSheet);
            Dictionary<string, List<DcSpec>> powerDcSpecs = isSplitDcSpecs ? dcGenerator.GetPowerDcSpecsByBlock() : dcGenerator.GetPowerDcSpecs();
            Dictionary<string, List<DcSpec>> ioDcSpecs = isSplitDcSpecs ? dcGenerator.GetIoDcSpecsByBlock() : dcGenerator.GetIoDcSpecs();

            var mergeSpecs = new Dictionary<string, List<DcSpec>>();
            var keys = powerDcSpecs.Select(x => x.Key).ToList();
            keys.AddRange(ioDcSpecs.Select(x => x.Key).ToList());
            foreach (string key in keys.Distinct())
            {
                if (powerDcSpecs.ContainsKey(key))
                {
                    if (mergeSpecs.ContainsKey(key))
                    {
                        mergeSpecs[key].AddRange(powerDcSpecs[key]);
                    }
                    else
                    {
                        mergeSpecs.Add(key, powerDcSpecs[key]);
                    }
                }

                if (ioDcSpecs.ContainsKey(key))
                {
                    if (mergeSpecs.ContainsKey(key))
                    {
                        mergeSpecs[key].AddRange(ioDcSpecs[key]);
                    }
                    else
                    {
                        mergeSpecs.Add(key, ioDcSpecs[key]);
                    }
                }
            }

            Mid.MultiDcSpecSheets = new List<DcSpecSheet>();
            foreach (KeyValuePair<string, List<DcSpec>> entry in mergeSpecs)
            {
                string sheetName = "DC_Specs_" + entry.Key;
                var selectorNameList = new List<string> { "Min", "Typ", "Max" };
                var list = entry.Value.First().CategoryList.Select(x => x.Name).ToList();
                var powerDcSpecSheet = new DcSpecSheet(sheetName, list, selectorNameList);
                powerDcSpecSheet.AddRows(entry.Value);
                Mid.MultiDcSpecSheets.Add(powerDcSpecSheet);
            }
        }

        protected virtual void GenLevel()
        {
            if (BasicResult.Level)
            {
                bool isSplitDc = MultiTestSettingSheetsSingleton.DcCategoryInfos.IsSplitDcSpecs(LocalSpecs.Options.IsSplitDcSpecs);
                Response.Report("Parsing Level Sets from Basic_Configure File ...", percentage: ProcessStatus = 55);
                var levelGenerator = new LevelSheetsGenerator(MultiTestSettingSheetsSingleton, TestPlanStatic.PowerInfoSheet, TestProgram.IgxlWorkBk.PinMapPair.Value);
                var level = new LevelInitial(TestPlanStatic.IoInfoSheet, TestPlanStatic.IoInfoConcurrentSheet, isSplitDc);
                Dictionary<string, LevelData> levelList = level.InitialLevelDatas();
                Mid.MultiLevelSheets = levelGenerator.GenerateFlow(levelList, TestPlanStatic.HardIpDcSheet, BasicInputData.IoContiSheet);
            }
        }

        protected virtual void GenGlobalSpec()
        {
            Response.Report("Generating Global_SPEC ...", percentage: ProcessStatus = 35);
            Mid.GlbSpecSheet = new GlobalSpecSheet("Global Specs");
            var globalGenerator = new GlobalSpecGenerator(BasicInputData.PinMapSheet, BasicInputData.IoContiSheet);
            List<GlobalSpec> globalSpecs = globalGenerator.GetPowerGlobalSpecs(TestPlanStatic.PowerInfoSheet, MultiTestSettingSheetsSingleton, BasicInputData.IfoldPowerTableSheet);
            List<GlobalSpec> ioGlobalSpecs = globalGenerator.GetIoGlobalSpecs(TestPlanStatic.IoInfoSheet, TestPlanStatic.IoInfoConcurrentSheet);
            Mid.GlbSpecSheet.AddRows(globalSpecs);
            Mid.GlbSpecSheet.AddRows(ioGlobalSpecs);
            //Add some special global specs
            var glbSymbolPlus = new GlbSymbolGenerator();
            glbSymbolPlus.DefaultGlbSymbols(Mid.GlbSpecSheet);
            glbSymbolPlus.PlusGlbSymbols(Mid.GlbSpecSheet);
            glbSymbolPlus.TestSettingPowerPinsShmooGlb(Mid.GlbSpecSheet, MultiTestSettingSheetsSingleton);

            ExcelWorksheet rtosTableSheet = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Find(x => x.Name.Equals("Rtos_Table", StringComparison.OrdinalIgnoreCase));
            if (rtosTableSheet != null)
            {
                var config = RtosTableSheet.LoadConfig(rtosTableSheet);
                Dictionary<string, string> rtosTablePins = config.PinRows;
                foreach (KeyValuePair<string, string> row in rtosTablePins)
                {
                    string pinName = row.Key;
                    string pinValue = row.Value;
                    GlobalSpec target = Mid.GlbSpecSheet.Rows.Find(x => x.Symbol.Equals(pinName, StringComparison.OrdinalIgnoreCase));
                    if (target != null)
                    {
                        target.Value = pinValue;
                    }
                    else
                    {
                        Mid.GlbSpecSheet.Rows.Add(new GlobalSpec(pinName, pinValue));
                    }
                }
            }
        }

        public void GenMappingTable()
        {
            ExcelWorksheet sheet = EpWorkbook.TestPlanWorkbook.Worksheets["SELSRM_Mapping_Table"] ?? EpWorkbook.TestPlanWorkbook.Worksheets["SELSRAM_Mapping_Table"];
            if (EpWorkbook.ScghWorkbook != null && sheet != null && string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                Response.Report("Generating DSSC MappingTable ...", percentage: 80);
                var dsscMappingTable = new NewDsscMappingTable();
                dsscMappingTable.Workflow(sheet, FolderStructure.DirCommonSheets, MultiTestSettingSheetsSingleton);
            }
        }

        private void GenNonIgxlSheet()
        {
            var specialTsValueFileGenerator = new SpecialTsValueFileGenerator(MultiTestSettingSheetsSingleton.TestSettingSheetsList, FolderStructure.DirCommonSheets);
            _specialTsValueFiles = specialTsValueFileGenerator.GenerateSpecialTsValueFile();
        }

        private void AddNonIgxlSheet()
        {
            if (_specialTsValueFiles != null && _specialTsValueFiles.Count > 0)
            {
                foreach (string file in _specialTsValueFiles)
                {
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, Path.GetFileNameWithoutExtension(file));
                }
            }
        }

        protected virtual void GenNwire()
        {
            if (!LocalSpecs.IsModuleIncluded("Basic"))
            {
                return;
            }

            if (NwireSingleton.Instance().HasNwirePin)
            {
                NwireResult = new NwireResult();

                var flow = new NwireFlow();
                NwireResult.NWireFlowSheets = flow.GenerateFlow();

                var instance = new NwireInstance();
                NwireResult.NWireInstanceRows = instance.GenerateFlow();

                var portMap = new NwirePortMap();
                NwireResult.PortMapSheets = portMap.GenerateFlow();

                var nwireBinTable = new NwireBintable();
                NwireResult.BinTables = nwireBinTable.GenerateFlow();
            }
        }

        protected virtual void GenRelay()
        {
            var relay = new RelayMain();
            LocalSpecs.RelayItems = null;
            if (EpWorkbook.TestPlanWorkbook.Worksheets["Relay"] != null)
            {
                //Create Relay information by new format relay on test plan
                LocalSpecs.RelayItems = relay.WorkFlow();

                EpWorkbook.TestPlanWorkbook.Worksheets["Relay"].ExportWorkBook2Txt(FolderStructure.DirCommon);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, "Relay");
            }
        }

        protected virtual void GenUfInstanceSheet()
        {
            if (TestPlanStatic.UfInstanceTable != null)
            {
                InstanceSheet testInstUfSheet = TestPlanStatic.UfInstanceTable.CreateInstanceSheet();
                foreach (InstanceRow row in testInstUfSheet.Rows)
                {
                    if (row.VbtType == ".NET")
                    {
                        row.VbtName = RenewFunctionName(row.VbtName);
                    }
                }
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, testInstUfSheet);
            }
        }

        protected virtual void GenContinuity()
        {
            if (BasicResult.Continuity)
            {
                ContiResult = new ContiResult();
                bool genConti = false;
                Response.Report("Generating continuity test ...", percentage: ProcessStatus = 90);

                List<BinTableRow> binTableRows = new List<BinTableRow>();
                foreach (DcTestContiSheet dcTestContiSheet in BasicInputData.DcTestContiSheets)
                {
                    var contiMain = new ContiMain(dcTestContiSheet, Mid.PatSetAll, TestPlanStatic.IoInfoSheet);
                    ContiResult result = contiMain.WorkFlow(Mid.MultiLevelSheets, ref binTableRows);
                    ContiResult.ContiResultList.Add(result);
                    genConti = true;
                }
                ContiResult.ContiBinTableRows = binTableRows;
                ContiResult = genConti ? ContiResult : null;
            }
        }

        private TimeSetSheets GetTimeSetSheets()
        {
            if (BasicInputData != null)
            {
                return BasicInputData.MultiTimeSetSheets;
            }

            TimeSetGenerator timeSetGenerator = new();
            TimeSetSheets timeSetSheets = timeSetGenerator.GenerateFlow(PatList, LocalSpecs.TimeSetFolder, FolderStructure.DirTimings);

            if (timeSetSheets.Count == 0)
            {
                Response.Report("TimeSet Path not existed, or found no TimeSet File!", EnumMessageLevel.Warning, ProcessStatus = 75);
            }

            if (timeSetGenerator.TimeSetWithWrongForamtRows.Any())
            {
                foreach (KeyValuePair<string, List<int>> timeSet in timeSetGenerator.TimeSetWithWrongForamtRows)
                {
                    Response.Report($"TSet Row Wrong Format: {timeSet.Key}, @RowNum: {string.Join(",", timeSet.Value)}, Compensate empty data", EnumMessageLevel.Error, ProcessStatus = 75);
                }
            }

            if (timeSetGenerator.TimeSetIncorrectFormat.Any())
            {
                foreach (string timeSet in timeSetGenerator.TimeSetIncorrectFormat)
                {
                    Response.Report($"TSet Row Wrong Format in {timeSet}, please check the timeSet.", EnumMessageLevel.Error, ProcessStatus = 75);
                }
            }

            return timeSetSheets;
        }

        protected virtual void GenTimeSet()
        {
            if (BasicResult.TimeSet)
            {
                Response.Report($"Copying Timing Set from path {LocalSpecs.TimeSetFolder} ...", percentage: ProcessStatus = 72);
                try
                {
                    Mid.TimeSetSheets = GetTimeSetSheets();

                    //Update MCG Mode
                    var mcgMode = new TimeSetMcgMode(NwireSingleton.Instance().SettingInfo.NwirePins);
                    mcgMode.ConverFlow(Mid.TimeSetSheets, PatList);

                    TimeSetGenerator generator = new();
                    TimeSetOverrideBatchApplier applier = new();
                    applier.Apply(Mid.TimeSetSheets, generator.GetTsetFileList(PatList))
                        .Report();

                    var timeSetPlus = new TimeSetPlus(NwireSingleton.Instance().NonFrcSetting);
                    timeSetPlus.PlusFlow(Mid.TimeSetSheets);

                    AcTSetCategoryMapSingleton.Instance().SetMultiTimeSetSheet(Mid.TimeSetSheets);

                }
                catch (Exception ex)
                {
                    Response.Report("Generating Timing Set failed! " + ex.Message, EnumMessageLevel.Warning, ProcessStatus = 90);
                }

            }
            else
            {
                Mid.TimeSetSheets = new TimeSetSheets();
                var timeSetPlus = new TimeSetPlus(NwireSingleton.Instance().NonFrcSetting);
                timeSetPlus.PlusFlow(Mid.TimeSetSheets);
            }
        }

        protected void GenPatternSet()
        {
            if (BasicResult.PattenSet)
            {
                Response.Report("Generating PatSet_All ...", percentage: ProcessStatus = 45);
                var patSetGenerator = new PatSetGenerator(LocalSpecs.PatternFolder);
                patSetGenerator.GenerateFlow(PatList);
                Mid.PatSetAll = patSetGenerator.PatSetSheetAll;
                Mid.PatSetSub = patSetGenerator.PatSubrSheetAll;
                patSetGenerator.GenPatSetErrorReport();
            }
        }

        protected void CopyNwireFile()
        {
            var ap = new NwireAp();
            ap.GenerateFlow(Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic"));
            var rf = new NwireRf();
            rf.GenerateFlow(Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic"));
            var lcd = new NwireLcd();
            lcd.GenerateFlow(Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic"));
        }

        #region At the End
        protected virtual void AddIgxlSheet()
        {
            if (Mid.GlbSpecSheet != null)
            {
                TestProgram.IgxlWorkBk.GlbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(FolderStructure.DirGlbSpec, Mid.GlbSpecSheet);
            }

            if (Mid.MultiDcSpecSheets != null)
            {
                foreach (DcSpecSheet dcSpecSheet in Mid.MultiDcSpecSheets)
                {
                    TestProgram.IgxlWorkBk.AddDcSpecSheet(FolderStructure.DirDcSpec, dcSpecSheet);
                }
            }

            if (Mid.AcSpecSheet != null)
            {
                TestProgram.IgxlWorkBk.AddAcSpecSheet(FolderStructure.DirAcSpec, Mid.AcSpecSheet);
            }

            if (Mid.MultiLevelSheets != null)
            {
                foreach (LevelSheet levelSheet in Mid.MultiLevelSheets.Values)
                {
                    TestProgram.IgxlWorkBk.AddLevelSheet(FolderStructure.DirLevel, levelSheet);
                }
            }

            if (Mid.TimeSetSheets != null)
            {
                foreach (ComTimeSetBasicSheet timeSetBasicSheet in Mid.TimeSetSheets)
                {
                    TestProgram.IgxlWorkBk.AddTimeSetSheet(FolderStructure.DirTimings, timeSetBasicSheet);
                }
            }

            if (Mid.PatSetAll != null)
            {
                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirPatSetsAll, Mid.PatSetAll);
            }

            if (Mid.PatSetSub != null)
            {
                TestProgram.IgxlWorkBk.AddPatSetSubSheet(FolderStructure.DirPatSetsAll, Mid.PatSetSub);
            }
            if (ContiResult != null)
            {
                foreach (ContiResult contiResult in ContiResult.ContiResultList)
                {
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirConti, contiResult.ContiFlowSheet);
                    TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirConti, contiResult.ConInstanceSheet);
                }
                BinTableSheet binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
                foreach (BinTableRow binTableRow in ContiResult.ContiBinTableRows)
                {
                    binTable.AddRow(binTableRow);
                }
                bool isExist = false;
                ContiResult.CommonInstance.Rows = new InstanceRows();
                ContiResult.CommonInstance.Rows.AddRange(ContiResult.ContiResultList.Select(p => p.CommonInstance).SelectMany(p => p.Rows).ToList());
                foreach (KeyValuePair<string, InstanceSheet> insSheetPair in TestProgram.IgxlWorkBk.InsSheets)
                {
                    if (insSheetPair.Value.Name == ContiResult.CommonInstance.Name)
                    {
                        isExist = true;
                        foreach (InstanceRow insRow in ContiResult.CommonInstance.Rows)
                        {
                            insSheetPair.Value.AddRow(insRow);
                        }
                    }
                }


                if (!isExist)
                {
                    TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, ContiResult.CommonInstance);
                }
            }

            if (NwireResult != null)
            {
                foreach (SubFlowSheet flowSheet in NwireResult.NWireFlowSheets)
                {
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirCommonSheets, flowSheet);
                }

                var inst = new InstanceSheet("TestInst_Nwire", "", true);
                inst.Rows.AddRange(NwireResult.NWireInstanceRows);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, inst);

                BinTableSheet binTableSheet = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
                foreach (BinTableRow binRow in NwireResult.BinTables)
                {
                    binTableSheet.AddRow(binRow);
                }
                foreach (PortMapSheet portMap in NwireResult.PortMapSheets)
                {
                    TestProgram.IgxlWorkBk.AddPortMapSheet(FolderStructure.DirPorts, portMap);
                }

            }

        }
        #endregion

        private string RenewFunctionName(string functionName)
        {
            return TestProgram.VbtFunctionLib.GetFunctionByName(functionName.Split('.').LastOrDefault(), "").FullFunctionName;
        }

        public void HarvestCheck()
        {
            HarvestCoreMappingChecker.HarvestCheck();
        }
    }
}
