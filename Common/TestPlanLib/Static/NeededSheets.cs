using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Enums;
using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.Mbist;
using TestPlanLib.Xml;

namespace TestPlanLib.Static
{
    public static class NeededSheets
    {
        public const string ParamMap = "VbtParamMapping";
        public const string VbtNameMap = "VbtNameMapping";
        public const string InstSpecialSetupMap = "InstSpecialSetting";
        public const string EnableWdGenerateTable = "EnableWdGenerateTable";

        public const string PrefixHardIp = "HARDIP_";
        public const string PrefixDctest = "DCTEST_";
        public const string PrefixRtos = "Rtos_";
        public const string PrefixWireless = "Wireless_";
        public const string PrefixLcd = "LCD_";

        public const string BinCutModeSequence = "BinCut Mode Sequence";
        public const string PostBinCutModeSeq = "Post BinCut Mode Sequence";
        public const string BcAte = "ATE Test Conditions";

        #region Property
        #region Test Plan
        private static string _testSettingTbl = "";
        private static string _vddLevels = "";
        private static string _ioLevels = "";
        private static string _powerTbl = "";
        private static string _powerGb = "";
        private static string _powerMerge = "";
        private static string _powerPerform = "";
        private static string _powerVmargin = "";
        private static string _contiDcTest = "";
        private static string _contiIo = "";
        private static string _ioGroup = "";
        private static string _pinMap = "";
        private static string _arraysSize = "";
        private static string _evsPat = "";
        private static string _evsFlow = "";
        private static string _mainFlow = "";
        private static string _hardIRegAssign = "";
        private static string _hardIPllmeas = "";
        private static string _hardIpDc = "";
        private static string _powerOverWrite = "";
        private static string _hardIpDcDefault = "";
        private static string _powerMemoryRepairBist = "";
        private static string _ioMapping = "";
        private static string _powerInfo = "";
        private static string _ioInfo = "";
        private static string _ioInfoConcurrent = "";
        private static string _ifoldPowerTable = "";
        private static string _clockpllmeas = "";
        private static string _pset = "";
        private static string _charSetting = "";

        public static string TestSettingTbl { get { return _testSettingTbl; } }
        public static string IoLevels { get { return _ioLevels; } }

        public static string PowerMerge { get { return _powerMerge; } }

        public static string ContiDcTest { get { return _contiDcTest; } }
        public static string ContiIo { get { return _contiIo; } }
        public static string IoGroup { get { return _ioGroup; } }
        public static string PinMap { get { return _pinMap; } }
        public static string ArraysSize { get { return _arraysSize; } }
        public static string MainFlow { get { return _mainFlow; } }
        public static string HardIRegAssign { get { return _hardIRegAssign; } }
        public static string HardIpPllMeas { get { return _hardIPllmeas; } }
        public static string HardIpDc { get { return _hardIpDc; } }
        public static string IoMapping { get { return _ioMapping; } }
        public static string ClockPllMeas { get { return _clockpllmeas; } }
        public static string PSet { get { return _pset; } }
        public static string CharSetting { get { return _charSetting; } }
        //New TestSetting Format
        public static string PowerInfo { get { return _powerInfo; } }
        public static string IoInfo { get { return _ioInfo; } }
        public static string IoInfoConcurrent { get { return _ioInfoConcurrent; } }
        public static string IfoldPwrTable { get { return _ifoldPowerTable; } }

        #endregion

        #region SCGH
        private static string _hardIpScgh = "";
        private static string _hardIpScghS = "";
        private static string _hardIpScghC = "";
        private static string _hardIpScghG = "";
        private static string _scanScgh = "";
        private static string _scanScghCpu = "";
        private static string _scanScghGpu = "";
        private static string _scanScghSoc = "";
        private static string _mbistScghCpu = "";
        private static string _mbistScghGpu = "";
        private static string _mbistScghSoc = "";
        private static string _mbistCharScgh = "";
        private static string _mbistCharScghCpu = "";
        private static string _mbistCharScghGpu = "";
        private static string _mbistCharScghSoc = "";
        private static string _newmbistScghCpu = "";
        private static string _newmbistScghGpu = "";
        private static string _newmbistScghSoc = "";
        private static string _spiScghCpu = "";
        private static string _spiScghGpu = "";
        private static string _spiScghSoc = "";
        private static string _pmicScgh = "";

        public static string HardIpScgh { get { return _hardIpScgh; } }
        public static string HardIpScghS { get { return _hardIpScghS; } }
        public static string HardIpScghC { get { return _hardIpScghC; } }
        public static string HardIpScghG { get { return _hardIpScghG; } }
        public static string ScanScgh { get { return _scanScgh; } }
        public static string ScanScghCpu { get { return _scanScghCpu; } }
        public static string ScanScghGpu { get { return _scanScghGpu; } }
        public static string ScanScghSoc { get { return _scanScghSoc; } }
        public static string MbistScghCpu { get { return _mbistScghCpu; } }
        public static string MbistScghGpu { get { return _mbistScghGpu; } }
        public static string MbistScghSoc { get { return _mbistScghSoc; } }
        public static string MbistCharScg { get { return _mbistCharScgh; } }
        public static string MbistCharScgCpu { get { return _mbistCharScghCpu; } }
        public static string MbistCharScgGpu { get { return _mbistCharScghGpu; } }
        public static string MbistCharScgSoc { get { return _mbistCharScghSoc; } }
        public static string SpiScghCpu { get { return _spiScghCpu; } }
        public static string SpiScghGpu { get { return _spiScghGpu; } }
        public static string SpiScghSoc { get { return _spiScghSoc; } }
        public static string PmicScgh { get { return _pmicScgh; } }
        public static List<MbistSheetInfo> MbistDomainOrder { get; private set; } = [];
        public static string RegexMbistSheets
        {
            get
            {
                var sheets = new List<string>();
                sheets.AddRange(MbistScghSoc.Split(','));
                sheets.AddRange(MbistScghCpu.Split(','));
                sheets.AddRange(MbistScghGpu.Split(','));
                return string.Join("|", sheets);
            }
        }

        public static string RegexMbistCharSheets
        {
            get
            {
                var sheets = new List<string>();
                sheets.AddRange(MbistCharScgCpu.Split(','));
                sheets.AddRange(MbistCharScgGpu.Split(','));
                sheets.AddRange(MbistCharScgSoc.Split(','));
                return string.Join("|", sheets);
            }
        }

        public static string RegexScanSheets
        {
            get
            {
                var sheets = new List<string>();
                sheets.AddRange(ScanScghCpu.Split(','));
                sheets.AddRange(ScanScghGpu.Split(','));
                sheets.AddRange(ScanScghSoc.Split(','));
                return string.Join("|", sheets);
            }
        }
        #endregion

        #region BinCut
        private static string _bcFlow = "";
        private static string _binning = "";
        private static string _binningBinX = "";
        private static string _binningBinY = "";
        private static string _binningNote = "";
        private static string _bcFlowPost = "";
        private static string _bcNotesPost = "";

        public static string BcFlow { get { return _bcFlow; } }
        public static string Binning { get { return _binning; } }
        public static string BinningBinX { get { return _binningBinX; } }
        public static string BinningBinY { get { return _binningBinY; } }
        public static string BinningNote { get { return _binningNote; } }

        public static string BcFlowPost { get { return _bcFlowPost; } }
        public static string BcNotesPost { get { return _bcNotesPost; } }
        #endregion

        #endregion

        #region Init Sheet name
        public static void InitSheetName(EnumDevice enumDevice, string settingFolder, string currentProject, string configFile = "")
        {
            string file = configFile;
            try
            {
                string needSheetFile;
                if (!string.IsNullOrEmpty(configFile))
                {
                    if (File.Exists(configFile))
                    {
                        needSheetFile = configFile;
                    }
                    else
                    {
                        throw new Exception($"{configFile} is not existed");
                    }
                }
                else
                {
                    needSheetFile = Path.Combine(settingFolder, "Settings", "Basic", $"NeedSheetConfig_{currentProject}.xml");
                }

                file = File.Exists(needSheetFile) ? needSheetFile : Path.Combine(settingFolder, "Settings", "Basic", "NeedSheetConfig_Default.xml");
                NeedSheetsConfig config = XmlService<NeedSheetsConfig>.LoadXml(file);
                if (enumDevice == EnumDevice.AP || enumDevice == EnumDevice.None)
                {
                    SetSheetName(ref _testSettingTbl, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("TestSettingTbl")));
                    SetSheetName(ref _vddLevels, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("VDDLevels")));
                    SetSheetName(ref _ioLevels, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IOLevels")));
                    SetSheetName(ref _powerTbl, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerTbl")));
                    SetSheetName(ref _powerGb, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerGb")));
                    SetSheetName(ref _powerMerge, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerMerge")));
                    SetSheetName(ref _powerPerform, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerPerform")));
                    SetSheetName(ref _powerVmargin, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerVmargin")));
                    SetSheetName(ref _contiDcTest, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ContiDcTest")));
                    SetSheetName(ref _contiIo, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ContiIo")));
                    SetSheetName(ref _ioGroup, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IoGroup")));
                    SetSheetName(ref _pinMap, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PinMap")));
                    SetSheetName(ref _arraysSize, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ArraysSize")));
                    SetSheetName(ref _evsPat, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("EvsPat")));
                    SetSheetName(ref _evsFlow, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("EvsFlow")));
                    SetSheetName(ref _mainFlow, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MainFlow")));
                    SetSheetName(ref _hardIPllmeas, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIPPLLMeas")));
                    SetSheetName(ref _hardIRegAssign, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIRegAssign")));
                    SetSheetName(ref _hardIpDc, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIPDc")));
                    SetSheetName(ref _powerOverWrite, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerOverWrite")));
                    SetSheetName(ref _hardIpDcDefault, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIpDcDefault")));
                    SetSheetName(ref _powerMemoryRepairBist, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerMemoryRepairBist")));
                    SetSheetName(ref _hardIpScgh, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIpScgh")));
                    SetSheetName(ref _hardIpScghS, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("S_HardIpScgh")));
                    SetSheetName(ref _hardIpScghC, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("C_HardIpScgh")));
                    SetSheetName(ref _hardIpScghG, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("G_HardIpScgh")));
                    SetSheetName(ref _scanScgh, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ScanScg")));
                    SetSheetName(ref _scanScghCpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ScanScgCpu")));
                    SetSheetName(ref _scanScghGpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ScanScgGpu")));
                    SetSheetName(ref _scanScghSoc, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ScanScgSoc")));
                    SetSheetName(ref _mbistScghCpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistScgCpu")));
                    SetSheetName(ref _mbistScghGpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistScgGpu")));
                    SetSheetName(ref _mbistScghSoc, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistScgSoc")));
                    SetSheetName(ref _mbistCharScgh, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistCharScg")));
                    SetSheetName(ref _mbistCharScghCpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistCharScgCpu")));
                    SetSheetName(ref _mbistCharScghGpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistCharScgGpu")));
                    SetSheetName(ref _mbistCharScghSoc, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistCharScgSoc")));
                    SetSheetName(ref _spiScghCpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("SpiScghCpu")));
                    SetSheetName(ref _spiScghGpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("SpiScghGpu")));
                    SetSheetName(ref _spiScghSoc, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("SpiScghSoc")));
                    SetSheetName(ref _pmicScgh, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PMICscgh")));
                    SetSheetName(ref _bcFlow, config.BinCutWorkBook.ToList().Find(p => p.ItemName.EqualsIgnoreCase("BcFlow")));
                    SetSheetName(ref _binning, config.BinCutWorkBook.ToList().Find(p => p.ItemName.EqualsIgnoreCase("Binning")));
                    SetSheetName(ref _binningBinX, config.BinCutWorkBook.ToList().Find(p => p.ItemName.EqualsIgnoreCase("BinningBinX")));
                    SetSheetName(ref _binningBinY, config.BinCutWorkBook.ToList().Find(p => p.ItemName.EqualsIgnoreCase("BinningBinY")));
                    SetSheetName(ref _binningNote, config.BinCutWorkBook.ToList().Find(p => p.ItemName.EqualsIgnoreCase("BcNotes")));
                    if (config.BinCutPostWorkbook != null)
                    {
                        SetSheetName(ref _bcFlowPost, config.BinCutPostWorkbook.ToList().Find(p => p.ItemName.EqualsIgnoreCase("BcFlowPost")));
                        SetSheetName(ref _bcNotesPost, config.BinCutPostWorkbook.ToList().Find(p => p.ItemName.EqualsIgnoreCase("BcNotesPost")));
                    }
                    SetSheetName(ref _ioMapping, config.IoMapping.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IoMapping")));
                    SetSheetName(ref _ioInfo, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IoInfo")));
                    SetSheetName(ref _ioInfoConcurrent, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IOInfo_Concurrent")));
                    SetSheetName(ref _powerInfo, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerInfo")));
                    SetSheetName(ref _ifoldPowerTable, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerTable")));
                    SetSheetName(ref _clockpllmeas, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ClockOutMeas")));
                    SetSheetName(ref _pset, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PSet")));
                    SetSheetName(ref _charSetting, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("CharSetting")));
                    GetMbistDomainOrder(config);
                }
                else if (enumDevice == EnumDevice.LCD || enumDevice == EnumDevice.RF)
                {
                    SetSheetName(ref _testSettingTbl, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("TestSettingTbl")));
                    SetSheetName(ref _vddLevels, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("VDDLevels")));
                    SetSheetName(ref _ioLevels, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IOLevels")));
                    SetSheetName(ref _powerTbl, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerTbl")));
                    SetSheetName(ref _powerGb, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerGb")));
                    SetSheetName(ref _powerMerge, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerMerge")));
                    SetSheetName(ref _powerPerform, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerPerform")));
                    SetSheetName(ref _powerVmargin, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerVmargin")));
                    SetSheetName(ref _contiDcTest, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ContiDcTest")));
                    SetSheetName(ref _contiIo, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ContiIo")));
                    SetSheetName(ref _ioGroup, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IoGroup")));
                    SetSheetName(ref _pinMap, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PinMap")));
                    SetSheetName(ref _arraysSize, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ArraysSize")));
                    SetSheetName(ref _evsPat, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("EvsPat")));
                    SetSheetName(ref _evsFlow, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("EvsFlow")));
                    SetSheetName(ref _mainFlow, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MainFlow")));
                    SetSheetName(ref _hardIPllmeas, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIPPLLMeas")));
                    SetSheetName(ref _hardIRegAssign, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIRegAssign")));
                    SetSheetName(ref _hardIpDc, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIPDc")));
                    SetSheetName(ref _powerOverWrite, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerOverWrite")));
                    SetSheetName(ref _hardIpDcDefault, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIpDcDefault")));
                    SetSheetName(ref _powerMemoryRepairBist, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerMemoryRepairBist")));
                    SetSheetName(ref _hardIpScgh, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("HardIpScgh")));
                    SetSheetName(ref _hardIpScghS, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("S_HardIpScgh")));
                    SetSheetName(ref _hardIpScghC, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("C_HardIpScgh")));
                    SetSheetName(ref _hardIpScghG, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("G_HardIpScgh")));
                    SetSheetName(ref _scanScgh, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ScanScg")));
                    SetSheetName(ref _scanScghCpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ScanScgCpu")));
                    SetSheetName(ref _scanScghGpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ScanScgGpu")));
                    SetSheetName(ref _scanScghSoc, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ScanScgSoc")));
                    SetSheetName(ref _mbistScghCpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistScgCpu")));
                    SetSheetName(ref _mbistScghGpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistScgGpu")));
                    SetSheetName(ref _mbistScghSoc, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistScgSoc")));
                    SetSheetName(ref _mbistCharScgh, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistCharScg")));
                    SetSheetName(ref _mbistCharScghCpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistCharScgCpu")));
                    SetSheetName(ref _mbistCharScghGpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistCharScgGpu")));
                    SetSheetName(ref _mbistCharScghSoc, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("MbistCharScgSoc")));
                    SetSheetName(ref _spiScghCpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("SpiScghCpu")));
                    SetSheetName(ref _spiScghGpu, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("SpiScghGpu")));
                    SetSheetName(ref _spiScghSoc, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("SpiScghSoc")));
                    SetSheetName(ref _pmicScgh, config.Scgh.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PMICscgh")));
                    SetSheetName(ref _ioMapping, config.IoMapping.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IoMapping")));
                    SetSheetName(ref _ioInfo, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IoInfo")));
                    SetSheetName(ref _ioInfoConcurrent, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("IOInfo_Concurrent")));
                    SetSheetName(ref _powerInfo, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerInfo")));
                    SetSheetName(ref _ifoldPowerTable, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PowerTable")));
                    SetSheetName(ref _clockpllmeas, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("ClockOutMeas")));
                    SetSheetName(ref _pset, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("PSet")));
                    SetSheetName(ref _charSetting, config.TestPlan.ToList().Find(p => p.ItemName.EqualsIgnoreCase("CharSetting")));
                    GetMbistDomainOrder(config);
                }
            }
            catch (Exception)
            {
                throw new Exception($"Please cheek if {file} is up to date !!!");
            }
        }

        private static void GetMbistDomainOrder(NeedSheetsConfig needSheetsConfig)
        {
            var mbistDomainOrder = new List<MbistSheetInfo>();
            int cpuIdx = needSheetsConfig.Scgh.ToList().FindIndex(x => x.ItemName.EqualsIgnoreCase("MbistScgCpu"));
            int gpuIdx = needSheetsConfig.Scgh.ToList().FindIndex(x => x.ItemName.EqualsIgnoreCase("MbistScgGpu"));
            int socIdx = needSheetsConfig.Scgh.ToList().FindIndex(x => x.ItemName.EqualsIgnoreCase("MbistScgSoc"));
            mbistDomainOrder.Add(new MbistSheetInfo { Domain = "cpu", DomainIdx = cpuIdx, Sheets = MbistScghCpu });
            mbistDomainOrder.Add(new MbistSheetInfo { Domain = "gpu", DomainIdx = gpuIdx, Sheets = MbistScghGpu });
            mbistDomainOrder.Add(new MbistSheetInfo { Domain = "soc", DomainIdx = socIdx, Sheets = MbistScghSoc });
            MbistDomainOrder = [.. mbistDomainOrder.OrderBy(x => x.DomainIdx)];
        }

        private static void SetSheetName(ref string value, NeedSheetItem? needSheetItem)
        {
            if (needSheetItem == null)
            {
                return;
            }

            value = needSheetItem.SheetName;

            if (needSheetItem.ItemName.EqualsIgnoreCase("MbistScgCpu"))
            {
                SetNewbist(ref _newmbistScghCpu, needSheetItem);
            }

            if (needSheetItem.ItemName.EqualsIgnoreCase("MbistScgSoc"))
            {
                SetNewbist(ref _newmbistScghSoc, needSheetItem);
            }

            if (needSheetItem.ItemName.EqualsIgnoreCase("MbistScgGpu"))
            {
                SetNewbist(ref _newmbistScghGpu, needSheetItem);
            }
        }

        private static void SetNewbist(ref string value, NeedSheetItem needSheetItem)
        {
            if (string.IsNullOrEmpty(needSheetItem.UseNewBist))
            {
                return;
            }

            if (needSheetItem.UseNewBist.EqualsIgnoreCase("true"))
            {
                value = needSheetItem.SheetName;
            }
        }
        #endregion

        #region Other Functions
        public static bool IsTestSettingSheetName(string sheetName, string currentProjectName)
        {
            var regTestSetting = new Regex(TestSettingTbl, RegexOptions.IgnoreCase);
            if (regTestSetting.IsMatch(sheetName))
            {
                string projectname = regTestSetting.Match(sheetName).Groups["projectname"].ToString();
                if (projectname.Length == 0 || projectname.EqualsIgnoreCase(currentProjectName))
                {
                    return true;
                }
            }
            return false;
        }

        public static string FindTestSettingSheetNameByJob(ExcelWorkbook excelWorkbook, string currentProjectName, string job)
        {
            List<ExcelWorksheet> alltestSettings = [.. excelWorkbook.Worksheets.Where(s => IsTestSettingSheetName(s.Name, currentProjectName))];
            ExcelWorksheet? jobTSsheet = alltestSettings.Find(s => Regex.Match(s.Name, TestSettingTbl, RegexOptions.IgnoreCase).Groups["job"].ToString().EqualsIgnoreCase(job)) ??
                                        alltestSettings.Find(s => Regex.Match(s.Name, TestSettingTbl, RegexOptions.IgnoreCase).Groups["job"].ToString().Length == 0);
            return jobTSsheet != null ? jobTSsheet.Name : "";
        }

        public static List<string> GetProdCharSheets()
        {
            var prodCharSheets = new List<string>();
            prodCharSheets.AddRange(_hardIpScgh.Split(','));
            prodCharSheets.AddRange(_hardIpScghS.Split(','));
            prodCharSheets.AddRange(_hardIpScghC.Split(','));
            prodCharSheets.AddRange(_hardIpScghG.Split(','));

            prodCharSheets.AddRange(_scanScgh.Split(','));
            prodCharSheets.AddRange(_scanScghCpu.Split(','));
            prodCharSheets.AddRange(_scanScghGpu.Split(','));
            prodCharSheets.AddRange(_scanScghSoc.Split(','));

            prodCharSheets.AddRange(_mbistCharScgh.Split(','));
            prodCharSheets.AddRange(_mbistCharScghCpu.Split(','));
            prodCharSheets.AddRange(_mbistCharScghGpu.Split(','));
            prodCharSheets.AddRange(_mbistCharScghSoc.Split(','));

            prodCharSheets.AddRange(_spiScghCpu.Split(','));
            prodCharSheets.AddRange(_spiScghGpu.Split(','));
            prodCharSheets.AddRange(_spiScghSoc.Split(','));
            return [.. prodCharSheets.Where(x => !string.IsNullOrEmpty(x))];
        }

        #endregion
    }
}
