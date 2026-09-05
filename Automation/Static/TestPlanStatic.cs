using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.PostAction.GenMainFlow.Business;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using CommonLib.Enums;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib;
using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;
using TestPlanLib.Concurrent;
using TestPlanLib.DataStruct;
using TestPlanLib.Efuse.Input;
using TestPlanLib.HardIpDc.BaseData;
using TestPlanLib.HardIpDc.Business;
using TestPlanLib.Harvest;
using TestPlanLib.Parser;
using TestPlanLib.PowerMerge;
using TestPlanLib.Scan;
using TestPlanLib.Static;

namespace Automation.Static
{
    public static class TestPlanStatic
    {
        private static IoInfoSheet _ioInfoSheet;
        public static IoInfoSheet IoInfoSheet
        {
            get
            {
                return _ioInfoSheet ?? (_ioInfoSheet = ReadSheetIfExists(NeededSheets.IoInfo, s => new IoInfoReader().ReadSheet(s)));
            }
            set
            {
                _ioInfoSheet = value;
            }
        }
        private static IoInfoSheet _ioInfoConcurrentSheet;
        public static IoInfoSheet IoInfoConcurrentSheet
        {
            get
            {
                return _ioInfoConcurrentSheet ?? (_ioInfoConcurrentSheet = ReadSheetIfExists(NeededSheets.IoInfoConcurrent, s => new IoInfoReader().ReadSheet(s)));
            }
            set
            {
                _ioInfoConcurrentSheet = value;
            }
        }
        private static SubprogramMappingSheet _subprogramMappingSheet;
        public static SubprogramMappingSheet SubprogramMappingSheet
        {
            get
            {
                return _subprogramMappingSheet ?? (_subprogramMappingSheet = ReadSheetIfExists("SubprogramMapping", s => new SubprogramMappingReader().ReadSheet(s)));
            }
            set
            {
                _subprogramMappingSheet = value;
            }
        }

        private static VreTestCaseTable _vreTestCaseTable;
        public static VreTestCaseTable VreTestCaseTable
        {
            get
            {
                return _vreTestCaseTable ?? (_vreTestCaseTable = ReadSheetIfExists("VRE_Test_Scenarios", s => new VreTestCaseTableReader().ReadSheet(s)));
            }
            set
            {
                _vreTestCaseTable = value;
            }
        }

        private static VreMbistLookupTable _vreMbistLookupTable;

        public static VreMbistLookupTable VreMbistLookupTable
        {
            get
            {
                return _vreMbistLookupTable ?? (_vreMbistLookupTable = ReadSheetIfExists("VRE_Mbist_Lookup", s => new VreMbistLookupTableReader().ReadSheet(s)));
            }
            set
            {
                _vreMbistLookupTable = value;
            }
        }
        private static TestNameWidthTable _testNameWidthTable;
        public static TestNameWidthTable TestNameWidthTable
        {
            get
            {
                return _testNameWidthTable ?? (_testNameWidthTable = ReadSheetIfExists("TestNameWidth", s => new TestNameWidthReader().ReadSheet(s)));
            }
        }

        private static VariableInitTable _variableInitTable;
        public static VariableInitTable VariableInitTable
        {
            get
            {
                return _variableInitTable ?? (_variableInitTable = ReadSheetIfExists("Variable_Initialize", s => new VariableInitTableReader().ReadSheet(s)));
            }
        }

        private static MappingHarvestingSheet _mappingHarvestingSheet;
        public static MappingHarvestingSheet MappingHarvestingSheet
        {
            get
            {
                return _mappingHarvestingSheet ?? (_mappingHarvestingSheet = ReadSheetRegex("Mapping_DigitalCores", s => new MappingHarvestingReader().ReadSheet(s)));
            }
        }

        public static MappingCoreTable MappingCoreTable { get; set; }

        private static HarvestingFlagInitSheet _harvestingFlagInitSheet;
        public static HarvestingFlagInitSheet HarvestingFlagInitSheet
        {
            get
            {
                return _harvestingFlagInitSheet ?? (_harvestingFlagInitSheet = ReadSheetStartsWith("HarvestingFlagInit", s => new HarvestingFlagInitReader(LocalSpecs.AllJobs).ReadSheet(s)));
            }
        }

        private static HarvestingFuseWriteSheet _harvestingFuseWriteSheet;
        public static HarvestingFuseWriteSheet HarvestingFuseWriteSheet
        {
            get
            {
                return _harvestingFuseWriteSheet ?? (_harvestingFuseWriteSheet = ReadSheetStartsWith("HarvestingFuseWrite", s => new HarvestingFuseWriteReader().ReadSheet(s)));
            }
        }

        private static ScanEnableWordSheet _scanEnableWordSheet;
        public static ScanEnableWordSheet ScanEnableWordSheet
        {
            get
            {
                return _scanEnableWordSheet ?? (_scanEnableWordSheet = ReadSheetRegex("Scan_Enable_words|Enable_words", s => new ScanEnableWordsReader().ReadSheet(s)));
            }
        }

        private static UfInstanceTable _ufInstanceTable;
        public static UfInstanceTable UfInstanceTable
        {
            get
            {
                return _ufInstanceTable ?? (_ufInstanceTable = ReadSheetIfExists("UF_Instance", s => new UfInstanceTableReader().ReadSheet(s)));
            }
        }

        private static UserFunctionSheet _userFunctionSheet;
        public static UserFunctionSheet UserFunctionSheet
        {
            get
            {
                return _userFunctionSheet ?? (_userFunctionSheet = ReadSheetIfExists("UserFunction", s => new UserFunctionReader().ReadSheet(s)));
            }
        }

        private static List<UfDigSrcSheet> _ufDigSrcSheets;
        public static List<UfDigSrcSheet> UfDigSrcSheets
        {
            get
            {
                if (_ufDigSrcSheets == null)
                {
                    ExcelWorkbook wb = EpWorkbook.TestPlanWorkbook;
                    if (wb == null)
                    {
                        return new List<UfDigSrcSheet>();
                    }

                    var sheets = wb.Worksheets.Where(s => s.Name.StartsWith("UF_DigSrc_", StringComparison.OrdinalIgnoreCase)).ToList();
                    _ufDigSrcSheets = new List<UfDigSrcSheet>();
                    foreach (ExcelWorksheet sheet in sheets)
                    {
                        _ufDigSrcSheets.AddRange(new UfDigSrcReader().ReadSheet(sheet));
                    }
                }
                return _ufDigSrcSheets;
            }
        }

        private static List<FlagOperationSheet> _flagOperationSheets;
        public static List<FlagOperationSheet> FlagOperationSheets
        {
            get
            {
                return _flagOperationSheets ?? (_flagOperationSheets = ReadSheetsStartsWith("FlagOperation", s => new FlagOperationReader().ReadSheet(s)));
            }
        }

        private static AteStrSummarySheet _ateStrSummarySheet;
        public static AteStrSummarySheet AteStrSummarySheet
        {
            get
            {
                return _ateStrSummarySheet ?? (_ateStrSummarySheet = ReadSheetIfExists("ATE_STR_Summary", s => new AteStrSummaryReader().ReadSheet(s)));
            }
        }

        private static List<DomainFlagsSheet> _domainFlagsSheets;
        public static List<DomainFlagsSheet> DomainFlagsSheets
        {
            get
            {
                return _domainFlagsSheets ?? (_domainFlagsSheets = ReadSheetsRegex("(SOC|CPU|GFX)_Flags", s => new DomainFlagsReader().ReadSheet(s)));
            }
        }

        internal static IdsMappingSheet _idsMappingSheet;
        public static IdsMappingSheet IdsMappingSheet
        {
            get
            {
                return _idsMappingSheet ?? (_idsMappingSheet = ReadSheetIfExists("IDS_Mapping_Table", s => new IdsMappingTableReader().ReadSheet(s)));
            }
        }

        private static PowerInfoSheet _powerInfoSheet;
        public static PowerInfoSheet PowerInfoSheet
        {
            get
            {
                return _powerInfoSheet ?? (_powerInfoSheet = ReadSheetIfExists(NeededSheets.PowerInfo, s => new PowerInfoReader().ReadSheet(s)));
            }
        }

        private static PowerMergeSheet _powerMergeSheet;
        public static PowerMergeSheet PowerMergeSheet
        {
            get
            {
                return _powerMergeSheet ?? (_powerMergeSheet = ReadSheetIfExists(NeededSheets.PowerMerge, PowerMergeReader.ReadSheet));
            }
            set
            {
                _powerMergeSheet = value;
            }
        }

        internal static List<HarvestingTruthTableSheet> _harvestingTruthTable;
        public static List<HarvestingTruthTableSheet> HarvestingTruthTableSheets
        {
            get
            {
                if (EpWorkbook.TestPlanWorkbook != null && _harvestingTruthTable == null)
                {
                    var harvestingTruthTables = new List<HarvestingTruthTableSheet>();
                    _harvestingTruthTable = harvestingTruthTables;
                    foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
                    {
                        if (Regex.IsMatch(sheet.Name, "^Harvesting_Truth[ |_]?Table_", RegexOptions.IgnoreCase))
                        {
                            var harvestingTruthTableReader = new HarvestingTruthTableReader(sheet.Name);
                            HarvestingTruthTableSheet harvTruthTableSheet = harvestingTruthTableReader.ReadSheet(sheet);
                            if (!harvTruthTableSheet.Check())
                            {
                                harvestingTruthTables.Add(harvTruthTableSheet);
                            }
                        }
                        else if (sheet.Name.StartsWith("HarvestingTruthTable", StringComparison.OrdinalIgnoreCase))
                        {
                            ExcelWorksheet efuseBitDefTable = EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_BitDef_Table"];
                            if (efuseBitDefTable == null)
                            {
                                return _harvestingTruthTable;
                            }

                            var harvestingTruthTableReader = new HarvestingTruthTableReader(sheet.Name);
                            HarvestingTruthTableSheet harvTruthTableSheet = harvestingTruthTableReader.ReadSheet(sheet);
                            harvTruthTableSheet.AddToErrorReport();
                            List<BitDefTable> efuseBitDefTables = EfuseBitDefTableReader.Read(efuseBitDefTable);
                            FuseWriteDataFilling(HarvestingFuseWriteSheet, efuseBitDefTables);
                            FlagInitBankFilling(HarvestingFlagInitSheet, efuseBitDefTables);
                            harvestingTruthTables.AddRange(new CreateHarvestTruthTable(harvTruthTableSheet, HarvestingFlagInitSheet, HarvestingFuseWriteSheet, LocalSpecs.AllJobs).Create());
                        }
                    }
                    TruthTableDataFilling(ref harvestingTruthTables);
                    _harvestingTruthTable = harvestingTruthTables;
                    foreach (HarvestingTruthTableSheet tableSheet in _harvestingTruthTable)
                    {
                        tableSheet.AddToErrorReport();
                    }
                }
                return _harvestingTruthTable;
            }
        }

        private static void FuseWriteDataFilling(HarvestingFuseWriteSheet harvestingFuseWriteSheet, List<BitDefTable> efuseBitDefTables)
        {
            if (harvestingFuseWriteSheet == null || efuseBitDefTables == null)
            {
                return;
            }

            foreach (HarvestingFuseWriteRow row in harvestingFuseWriteSheet.Rows)
            {
                string fieldName;
                string blockName = "";
                string[] items = row.FuseName.Split('.');
                if (items.Length > 1)
                {
                    blockName = items[0].ToLower().Replace("bank_", "");
                    fieldName = items[1].Split('[', ':')[0];
                }
                else
                {
                    fieldName = items[0];
                }
                foreach (BitDefTable bitDefTable in efuseBitDefTables)
                {
                    int fuseNameIdx = bitDefTable.BankEfuseBitDefIdx;
                    int stageIdx = bitDefTable.ProgrammingStageIdx;
                    if (!string.IsNullOrEmpty(blockName) && !bitDefTable.BlockName.Equals(blockName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    BitDefRow target = bitDefTable.Rows.Find(x => x.RowData[fuseNameIdx].Equals(fieldName, StringComparison.OrdinalIgnoreCase));
                    if (target != null)
                    {
                        row.Job = target.RowData[stageIdx];
                        row.BlockName = bitDefTable.BlockName;
                        break;
                    }
                }
            }
        }

        private static void FlagInitBankFilling(HarvestingFlagInitSheet harvestingFlagInitSheet, List<BitDefTable> efuseBitDefTables)
        {
            if (harvestingFlagInitSheet == null || efuseBitDefTables == null)
            {
                return;
            }

            foreach (string fuseName in harvestingFlagInitSheet.EfuseMappingTable.SelectMany(x => x.FuseMapping).Select(y => y.Value).Distinct())
            {
                if (harvestingFlagInitSheet.BankDictionary.ContainsKey(fuseName))
                {
                    continue;
                }

                foreach (BitDefTable bitDefTable in efuseBitDefTables)
                {
                    int fuseNameIdx = bitDefTable.BankEfuseBitDefIdx;
                    BitDefRow target = bitDefTable.Rows.Find(x => x.RowData[fuseNameIdx].Equals(fuseName.Split('[').FirstOrDefault(), StringComparison.OrdinalIgnoreCase));
                    if (target != null)
                    {
                        harvestingFlagInitSheet.BankDictionary.Add(fuseName, bitDefTable.BlockName);
                    }
                }
            }
        }

        private static void TruthTableDataFilling(ref List<HarvestingTruthTableSheet> harvestingTruthTables)
        {
            if (harvestingTruthTables != null)
            {
                var allFlags =
                    harvestingTruthTables.SelectMany(x => x.Rows)
                        .Select(x => x.Flags)
                        .SelectMany(x => x)
                        .ToList();
                var mergeValueByFlag = new Dictionary<string, List<string>>();
                foreach (Flag flag in allFlags)
                {
                    if (flag.Value.Equals("X", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string flagName = flag.FlagName;
                    string flagValue = flag.Value;
                    if (mergeValueByFlag.ContainsKey(flagName))
                    {
                        if (!mergeValueByFlag[flagName].Contains(flagValue))
                        {
                            mergeValueByFlag[flagName].Add(flag.Value);
                        }
                    }
                    else
                    {
                        mergeValueByFlag.Add(flagName, new List<string> { flagValue });
                    }
                }
                foreach (HarvestingTruthTableSheet harvestingTruthTable in harvestingTruthTables)
                {
                    harvestingTruthTable.MergeValueByFlag = mergeValueByFlag;
                    if (harvestingTruthTable.Job.Equals("CP1"))
                    {
                        continue;
                    }

                    harvestingTruthTable.ReadFuseChkList = harvestingTruthTable.GetHarvestEfuseRead().ToList();
                    harvestingTruthTable.CheckEfuseRead();
                }
            }
        }

        internal static JobInfoSheet _jobInfoSheet;
        public static JobInfoSheet JobInfoSheet
        {
            get
            {
                return _jobInfoSheet ?? (_jobInfoSheet = ReadSheetIfExists("Job_Info", s => new JobInfoReader(LocalSpecs.AllJobs, LocalSpecs.Options.VreEnable).ReadSheet(s)));
            }
        }

        private static List<EnumEquipment> _equipments;
        public static List<EnumEquipment> Equipments
        {
            get
            {
                if (_equipments == null)
                {
                    if (JobInfoSheet != null)
                    {
                        var equipments = new List<EnumEquipment>();
                        IEnumerable<string> allTesterTypes = JobInfoSheet.Rows.Select(x => x.TesterType).Distinct();
                        foreach (string type in allTesterTypes)
                        {
                            if (type.ToUpper().Equals("UF"))
                            {
                                equipments.Add(EnumEquipment.UltraFlex);
                            }
                            else if (type.ToUpper().Equals("UFP"))
                            {
                                equipments.Add(EnumEquipment.UltraFlexPlus);
                            }
                        }

                        _equipments = equipments;
                    }
                    else
                    {
                        var equipments = new List<EnumEquipment>();
                        switch (LocalSpecs.Options.InstrumentType)
                        {
                            case EnumInstrument.UFlex_UFPlus:
                                equipments.Add(EnumEquipment.UltraFlex);
                                equipments.Add(EnumEquipment.UltraFlexPlus);
                                break;

                            case EnumInstrument.UFPlus:
                                equipments.Add(EnumEquipment.UltraFlexPlus);
                                break;

                            case EnumInstrument.UFlex:
                                equipments.Add(EnumEquipment.UltraFlex);
                                break;

                            default:
                                equipments.Add(EnumEquipment.UltraFlexPlus);
                                break;
                        }
                        _equipments = equipments;
                    }
                }
                return _equipments;
            }
        }

        private static HardIpDcSheet _hardIpDcSheet;
        public static HardIpDcSheet HardIpDcSheet
        {
            get
            {
                return _hardIpDcSheet ?? (_hardIpDcSheet = ReadSheetIfExists(NeededSheets.HardIpDc, s => new HardIpDcReader().ReadSheet(s)));
            }
        }

        private static List<BitDefTable> _bitDefTables;
        public static List<BitDefTable> BitDefTables
        {
            get
            {
                return _bitDefTables ?? (_bitDefTables = ReadSheetIfExists(SheetConst.Type5BitDefTable, EfuseBitDefTableReader.Read));
            }
        }

        private static List<BinCutInstanceSheet> _efuseInstanceSheets;
        public static List<BinCutInstanceSheet> EfuseInstanceSheets
        {
            get
            {
                return _efuseInstanceSheets ?? (_efuseInstanceSheets = ReadSheetsStartsWith("Instance_Efuse", s => new EfuseInstanceSheetReader().ReadSheet(s)));
            }
        }

        private static List<BinCutInstanceSheet> _binCutInstanceSheets;
        public static List<BinCutInstanceSheet> BinCutInstanceSheets
        {
            get
            {
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.BinCut).Down && _binCutInstanceSheets == null && EpWorkbook.TestPlanWorkbook != null)
                {
                    var binCutInstanceSheets = new List<BinCutInstanceSheet>();
                    foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
                    {
                        if ((sheet.Name.Equals("BinCut_Instance", StringComparison.CurrentCultureIgnoreCase) || sheet.Name.StartsWith("Instance_BinCut_", StringComparison.CurrentCultureIgnoreCase)) && !sheet.Name.ContainsIgnoreCase("post"))
                        {
                            Response.Report($"Reading {sheet.Name} ...", percentage: 20);
                            var binCutInstanceSheetReader = new BinCutInstanceSheetReader();
                            BinCutInstanceSheet binCutInstanceSheet = binCutInstanceSheetReader.ReadSheet(sheet);
                            if (binCutInstanceSheet != null)
                            {
                                binCutInstanceSheets.Add(binCutInstanceSheet);
                            }
                        }
                    }
                    _binCutInstanceSheets = binCutInstanceSheets;
                }
                return _binCutInstanceSheets;
            }
        }

        private static List<BinCutInstanceSheet> _htolInstanceSheets;
        public static List<BinCutInstanceSheet> HtolInstanceSheets
        {
            get
            {
                if (_htolInstanceSheets == null && EpWorkbook.TestPlanWorkbook != null)
                {
                    var htolInstanceSheets = new List<BinCutInstanceSheet>();
                    foreach (ExcelWorksheet worksheet in EpWorkbook.TestPlanWorkbook.Worksheets)
                    {
                        if (worksheet.Name.StartsWith("Instance_", StringComparison.CurrentCultureIgnoreCase) && worksheet.Name.ContainsIgnoreCase("htol"))
                        {
                            Response.Report($"Reading {worksheet.Name} ...", percentage: 20);
                            var htolInstanceSheetReader = new NonBinCutInstanceSheetReader();
                            BinCutInstanceSheet htolInstanceSheet = htolInstanceSheetReader.ReadSheet(worksheet);
                            htolInstanceSheets.Add(htolInstanceSheet);
                        }
                    }
                    _htolInstanceSheets = htolInstanceSheets;
                }
                return _htolInstanceSheets;
            }
        }

        private static List<BinCutInstanceSheet> _scanInstanceSheets;
        public static List<BinCutInstanceSheet> ScanInstanceSheets
        {
            get
            {
                if (_scanInstanceSheets == null && EpWorkbook.TestPlanWorkbook != null)
                {
                    var scanInstanceSheets = new List<BinCutInstanceSheet>();
                    foreach (ExcelWorksheet worksheet in EpWorkbook.TestPlanWorkbook.Worksheets)
                    {
                        if (worksheet.Name.StartsWith("BinCut_Instance", StringComparison.CurrentCultureIgnoreCase) || worksheet.Name.StartsWith("Instance_BinCut_", StringComparison.CurrentCultureIgnoreCase) || worksheet.Name.StartsWith("Instance_Post_BinCut", StringComparison.CurrentCultureIgnoreCase) || worksheet.Name.StartsWith("Instance_Efuse", StringComparison.CurrentCultureIgnoreCase) || worksheet.Name.ContainsIgnoreCase("clock") || worksheet.Name.ContainsIgnoreCase("evs") || worksheet.Name.ContainsIgnoreCase("htol") || worksheet.Name.ContainsIgnoreCase("cpm"))
                        {
                            continue;
                        }

                        if (worksheet.Name.StartsWith("Instance_", StringComparison.CurrentCultureIgnoreCase))
                        {
                            Response.Report($"Reading {worksheet.Name} ...", percentage: 20);
                            var scanInstanceSheetReader = new NonBinCutInstanceSheetReader();
                            BinCutInstanceSheet scanInstanceSheet = scanInstanceSheetReader.ReadSheet(worksheet);
                            if (scanInstanceSheet != null)
                            {
                                scanInstanceSheets.Add(scanInstanceSheet);
                            }
                        }
                    }
                    _scanInstanceSheets = scanInstanceSheets;
                }
                return _scanInstanceSheets;
            }
        }

        private static TimeSettingSheet _timeSettingSheet;
        public static TimeSettingSheet TimeSettingSheet
        {
            get
            {
                if (BlockStatus.GetAutomationBlockStatus(BlockConst.Basic).Down && _timeSettingSheet == null && EpWorkbook.TestPlanWorkbook != null)
                {
                    foreach (ExcelWorksheet worksheet in EpWorkbook.TestPlanWorkbook.Worksheets)
                    {
                        if (worksheet.Name.StartsWith("TimeSetting", StringComparison.CurrentCultureIgnoreCase))
                        {
                            Response.Report($"Reading {worksheet.Name} ...", percentage: 20);
                            var timeSettingReader = new TimeSettingReader();
                            TimeSettingSheet timeSettingSheet = timeSettingReader.ReadSheet(worksheet);
                            if (timeSettingSheet != null)
                            {
                                _timeSettingSheet = timeSettingSheet;
                            }
                        }
                    }
                }
                return _timeSettingSheet;
            }
        }

        private static List<BinCutInstanceSheet> _evsInstanceSheets;
        public static List<BinCutInstanceSheet> EvsInstanceSheets
        {
            get
            {
                if (_evsInstanceSheets == null && EpWorkbook.TestPlanWorkbook != null)
                {
                    var evsInstanceSheets = new List<BinCutInstanceSheet>();
                    foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
                    {
                        if (sheet.Name.StartsWith("Instance_EVS", StringComparison.CurrentCultureIgnoreCase))
                        {
                            Response.Report($"Reading {sheet.Name} ...", percentage: 20);
                            var evsInstanceSheetReader = new EvsInstanceSheetReader();
                            BinCutInstanceSheet evsInstanceSheet = evsInstanceSheetReader.ReadSheet(sheet);
                            if (evsInstanceSheet != null)
                            {
                                evsInstanceSheets.Add(evsInstanceSheet);
                            }
                        }

                    }
                    _evsInstanceSheets = evsInstanceSheets;
                }
                return _evsInstanceSheets;
            }
        }
        private static HardIpInputData _hardIpInputDataRtos;
        public static HardIpInputData RtosSheets
        {
            get
            {
                if (_hardIpInputDataRtos == null && EpWorkbook.TestPlanWorkbook != null && LocalSpecs.IsModuleIncluded(BlockStatus.Rtos))
                {
                    var hardIpParaData = new HardIpParaData(EnumBlock.Rtos);
                    _hardIpInputDataRtos = new HardIpInputManager(EpWorkbook.TestPlanWorkbook, hardIpParaData).Read();

                }
                return _hardIpInputDataRtos;
            }
        }

        private static HardIpInputData _hardIpInputDataIds;
        public static HardIpInputData IdsSheets
        {
            get
            {
                if (_hardIpInputDataIds == null && EpWorkbook.TestPlanWorkbook != null && LocalSpecs.IsModuleIncluded(BlockStatus.Ids))
                {
                    var hardIpParaData = new HardIpParaData(EnumBlock.Ids);
                    _hardIpInputDataIds = new HardIpInputManager(EpWorkbook.TestPlanWorkbook, hardIpParaData).Read();

                }
                return _hardIpInputDataIds;
            }
        }


        private static MainFlowSheet _mainFlowSheet;
        public static MainFlowSheet MainFlowSheet
        {
            get
            {
                if (_mainFlowSheet == null && EpWorkbook.TestPlanWorkbook != null && EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Exists(x => x.Name.Equals(NeededSheets.MainFlow)))
                {
                    ExcelWorksheet mainFlowWorksheet = EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.MainFlow];
                    if (mainFlowWorksheet != null)
                    {
                        var mainFlowSheetReader = new MainFlowSheetReader();
                        MainFlowSheet mainFlowSheet = mainFlowSheetReader.ReadSheet(mainFlowWorksheet);
                        if (mainFlowSheet.Rows.Any())
                        {
                            _mainFlowSheet = mainFlowSheet;
                        }
                        else
                        {
                            var mainFlowSheetReaderNew = new MainFlowSheetReaderNew();
                            mainFlowSheet = mainFlowSheetReaderNew.ReadSheet(mainFlowWorksheet);
                            _mainFlowSheet = mainFlowSheet;
                            LocalSpecs.SetEnableModules(mainFlowSheetReaderNew.GetEnableModules());
                        }
                    }
                }
                return _mainFlowSheet;
            }
            set
            {
                _mainFlowSheet = value;
            }
        }

        private static TestProgramDefSheet _testProgramDefSheet;
        public static TestProgramDefSheet TestProgramDefSheet
        {
            get
            {
                if (_testProgramDefSheet == null && EpWorkbook.TestPlanWorkbook != null && EpWorkbook.TestPlanWorkbook.Worksheets["TestProgram_Def"] != null)
                {
                    _testProgramDefSheet = new TestProgramDefReader().ReadSheet(EpWorkbook.TestPlanWorkbook.Worksheets["TestProgram_Def"]);
                }
                return _testProgramDefSheet;
            }
            set
            {
                _testProgramDefSheet = value;
            }
        }

        private static ConcurrentFlowSheet _concurrentFlowSheet;
        public static ConcurrentFlowSheet ConcurrentFlowSheet
        {
            get
            {
                return _concurrentFlowSheet ?? (_concurrentFlowSheet = ReadSheetIfExists("Concurrent_Flow", s => new ConcurrentFlowSheetReader().ReadSheet(s)));
            }
            set
            {
                _concurrentFlowSheet = value;
            }
        }

        private static List<BinCutInstanceSheet> _cpmInstanceSheets;
        public static List<BinCutInstanceSheet> CpmInstanceSheets
        {
            get
            {
                if (EpWorkbook.TestPlanWorkbook != null && _cpmInstanceSheets == null)
                {
                    var cpmInstanceSheets = new List<BinCutInstanceSheet>();
                    foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
                    {
                        if (sheet.Name.StartsWith("Instance_CPM", StringComparison.CurrentCultureIgnoreCase))
                        {
                            Response.Report($"Reading {sheet.Name} ...", percentage: 20);
                            var cpmInstanceSheetReader = new NonBinCutInstanceSheetReader();
                            BinCutInstanceSheet cpmInstanceSheet = cpmInstanceSheetReader.ReadSheet(sheet);
                            if (cpmInstanceSheet != null)
                            {
                                cpmInstanceSheets.Add(cpmInstanceSheet);
                            }
                        }
                    }
                    _cpmInstanceSheets = cpmInstanceSheets;
                }
                return _cpmInstanceSheets;
            }
        }

        private static DigitalFlagsSheet _digitalFlagsSheet;
        public static DigitalFlagsSheet DigitalFlagsSheet
        {
            get
            {
                return _digitalFlagsSheet ?? (_digitalFlagsSheet = ReadSheetIfExists("Digital_Flags", s => new DigitalFlagsReader().ReadSheet(s)));
            }
            set
            {
                _digitalFlagsSheet = value;
            }
        }

        private static MiscInfoParser _miscInfoParser;

        public static MiscInfoParser MiscInfoParser
        {
            get
            {
                return _miscInfoParser ?? (_miscInfoParser = new MiscInfoParser());
            }
        }

        private static TSheet ReadSheetIfExists<TSheet>(string sheetName, Func<ExcelWorksheet, TSheet> reader)
        {
            ExcelWorkbook wb = EpWorkbook.TestPlanWorkbook;
            if (wb == null)
            {
                return default;
            }

            ExcelWorksheet sheet = wb.Worksheets.FirstOrDefault(s => s.Name.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase));
            return sheet != null ? reader(sheet) : default;
        }

        private static TSheet ReadSheetStartsWith<TSheet>(string prefix, Func<ExcelWorksheet, TSheet> reader)
        {
            ExcelWorkbook wb = EpWorkbook.TestPlanWorkbook;
            if (wb == null)
            {
                return default;
            }

            ExcelWorksheet sheet = wb.Worksheets.FirstOrDefault(s =>
                s.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
            return sheet != null ? reader(sheet) : default;
        }

        private static List<TSheet> ReadSheetsStartsWith<TSheet>(string prefix, Func<ExcelWorksheet, TSheet> reader)
        {
            ExcelWorkbook wb = EpWorkbook.TestPlanWorkbook;
            if (wb == null)
            {
                return new List<TSheet>();
            }

            return wb.Worksheets.Where(s => s.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).Select(reader).ToList();
        }

        private static TSheet ReadSheetRegex<TSheet>(string pattern, Func<ExcelWorksheet, TSheet> reader)
        {
            ExcelWorkbook wb = EpWorkbook.TestPlanWorkbook;
            if (wb == null)
            {
                return default;
            }

            ExcelWorksheet sheet = wb.Worksheets.FirstOrDefault(s =>
                Regex.IsMatch(s.Name, pattern, RegexOptions.IgnoreCase));
            return sheet != null ? reader(sheet) : default;
        }

        private static List<TSheet> ReadSheetsRegex<TSheet>(string pattern, Func<ExcelWorksheet, TSheet> reader)
        {
            ExcelWorkbook wb = EpWorkbook.TestPlanWorkbook;
            if (wb == null)
            {
                return new List<TSheet>();
            }

            return wb.Worksheets
                .Where(s => Regex.IsMatch(s.Name, pattern, RegexOptions.IgnoreCase))
                .Select(reader)
                .ToList();
        }

        public static void Clear()
        {
            _ioInfoSheet = null;
            _ioInfoConcurrentSheet = null;
            _subprogramMappingSheet = null;
            _testNameWidthTable = null;
            _variableInitTable = null;
            _mappingHarvestingSheet = null;
            _harvestingFlagInitSheet = null;
            _harvestingFuseWriteSheet = null;
            _scanEnableWordSheet = null;
            _ufInstanceTable = null;
            _userFunctionSheet = null;
            _ufDigSrcSheets = null;
            _flagOperationSheets = null;
            _ateStrSummarySheet = null;
            _domainFlagsSheets = null;
            _idsMappingSheet = null;
            _powerInfoSheet = null;
            _powerMergeSheet = null;
            _harvestingTruthTable = null;
            MappingCoreTable = null;
            _jobInfoSheet = null;
            _equipments = null;
            _hardIpDcSheet = null;
            _bitDefTables = null;
            _efuseInstanceSheets = null;
            _binCutInstanceSheets = null;
            _htolInstanceSheets = null;
            _scanInstanceSheets = null;
            _timeSettingSheet = null;
            _evsInstanceSheets = null;
            _mainFlowSheet = null;
            _testProgramDefSheet = null;
            _concurrentFlowSheet = null;
            _cpmInstanceSheets = null;
            _digitalFlagsSheet = null;
            _miscInfoParser = null;
            _hardIpInputDataRtos = null;
            _hardIpInputDataIds = null;
        }
    }
}
