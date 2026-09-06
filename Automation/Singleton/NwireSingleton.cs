using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

using System.Text.RegularExpressions;
using System.Xml;

using Automation.Const;
using Automation.Reader;
using Automation.Static;
using Automation.Utility.Basic;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility.FrcCalc;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility.CurrentChannel;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.DataStruct;
using TestPlanLib.Settings;
using TestPlanLib.Static;

using SbcFreqCalculator = Automation.Utility.SbcFreqCalculator;

namespace Automation.Singleton
{
    public class NwireSingleton
    {
        private static NwireSingleton _instance;
        private static readonly object _locker = new object();
        private const string RegexPattern = @"(?<value>\d+([.]\d+)?)\s*(?<unit>\w*)";
        private static readonly Regex _regex = new Regex(@"t\d+:.*", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex4 = new Regex(CommonConst.PinRegPattern, RegexOptions.Compiled);
        private static readonly Regex _regex7 = new Regex(@"\d+\.ch\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex5 = new Regex(RegexPattern, RegexOptions.Compiled);

        public bool HasNwirePin { get; }
        public NwireSetting SettingInfo { set; get; } = new NwireSetting();
        public List<NonFrcNWires> NonFrcSetting { set; get; }
        public XmlNode Node { set; get; }          //LocalSpecs.SettingFolder + FRCRef.pa
        public XmlNode NodeDiff { set; get; }      //LocalSpecs.SettingFolder + FRCRef_differential.pa
        public XmlNode NodeAdg1414 { set; get; }   //LocalSpecs.SettingFolder + SPI.xml

        public static NwireSingleton Instance()
        {
            if (_instance == null)
            {
                lock (_locker)
                {
                    if (_instance == null)
                    {
                        _instance = new NwireSingleton();
                    }
                }
            }
            return _instance;
        }

        public static void Initialize()
        {
            _instance = null;
        }

        public const string NwireFlag = "F_nWire_lock_check";
        public const string SbcGlbSpecName = "SBC_Freq_Glb";

        internal NwireSingleton()
        {
            ExcelWorksheet worksheet = GetFreeRunClkSheet();
            if (worksheet != null)
            {
                var reader = new NwireReader();
                SettingInfo = reader.ReadFlow(worksheet);
                AddFrcPinFromPowerInfo(worksheet);

                if (SettingInfo.NwirePins.Any())
                {
                    HasNwirePin = true;
                }

                //Read the IO voltage of Out clock pin
                SetIoVoltage();

                var calculator = new SbcFreqCalculator();
                foreach (ProtocolAwarePin nWirePin in SettingInfo.NwirePins)
                {
                    //Can not find correct voltage of this pin
                    if (nWirePin.OutClkVoltage.Equals(0))
                    {
                        //nWirePin.OutClkVoltage = -2;
                        ErrorReportManager.AddError(BasicErrorType.W_NwireConfigMismatch_01, NeededSheets.ContiIo, 1, 1, $"Please check {nWirePin.OutClk} !!! Pin OutClkVoltage value is 0", new string[] { nWirePin.OutClk });
                    }
                    nWirePin.RefClk = string.IsNullOrEmpty(nWirePin.RefClk) ? GetRefClkName(nWirePin) : nWirePin.RefClk;
                    calculator.TargetFreq.Add((int)nWirePin.Freq);
                }

                if (TestPlanStatic.UfInstanceTable?.CheckExist("StartSBClock") == true)
                {
                    UfInstanceRow ufSbcInstance =
                        TestPlanStatic.UfInstanceTable.Rows
                            .FirstOrDefault(x => x.TestName.Equals("StartSBClock", StringComparison.CurrentCultureIgnoreCase));
                    int freqIndex = (ufSbcInstance.ArgList?.Split(',').ToList() ?? new List<string>())
                        .FindIndex(s => s.Equals("frequency", StringComparison.OrdinalIgnoreCase));

                    string item = ufSbcInstance.Arg.ElementAtOrDefault(freqIndex);
                    if (item != null)
                    {
                        if (double.TryParse(item, out double freqValue))
                        {
                            SettingInfo.SupportBoardFreq = freqValue;
                        }
                    }
                }
                else if (LocalSpecs.Options.FrcCalcType == "RF")
                {
                    List<double> sbc = FrcCalcMain.CalculateFrcFreq(calculator.TargetFreq);
                    SettingInfo.SupportBoardFreq = sbc[0];
                }
                else
                {
                    SettingInfo.SupportBoardFreq = calculator.SolveSbcFreq().SbcFreq;
                }
            }

        }

        private ExcelWorksheet GetFreeRunClkSheet()
        {
            #region RF_USED
            if (EpWorkbook.FreeRunningClockAndRelayFormatWorkbook == null && LocalSpecs.IsBenchLog)
            {
                string nWireRlySetting = LocalSpecs.SettingFolder + string.Format("\\Settings\\Basic\\nWire_and_relay_{0}.xlsx", LocalSpecs.CurrentProject);
                if (File.Exists(nWireRlySetting))
                {
                    var inputExcel = new ExcelPackage(new FileInfo(nWireRlySetting));
                    EpWorkbook.FreeRunningClockAndRelayFormatWorkbook = inputExcel.Workbook;
                }
                else
                {
                    return null;
                    //throw new Exception(string.Format("Can not find worksheet:{0} in nWire setting file", NeededSheets.FreeRunClk));
                }
            }
            #endregion
            if (EpWorkbook.TestPlanWorkbook != null)
            {
                ExcelWorksheet worksheet = EpWorkbook.TestPlanWorkbook.Worksheets["FreeRunningClock"];
                if (worksheet != null)
                {
                    return worksheet;
                }
            }
            return null;
        }

        private string GetRefClkName(ProtocolAwarePin nWirePin)
        {
            string refClk;
            string diffPinName = nWirePin.OutClk + "::" + nWirePin.OutClkDiff;
            DifferentialService.DiffPinPosAndNeg(diffPinName, out string _, out string _, out string groupName);
            if (groupName != "")
            {
                refClk = ProtocolAwarePin.ConRefClk + "_" + groupName;
            }
            else
            {
                refClk = ProtocolAwarePin.ConRefClk + "_" + nWirePin.OutClk + (!string.IsNullOrEmpty(nWirePin.OutClkDiff) ? "_" + nWirePin.OutClkDiff : "");
            }

            return refClk;
        }

        internal void AddFrcPinFromPowerInfo(ExcelWorksheet nWireWorksheet)
        {
            if (EpWorkbook.TestPlanWorkbook == null)
            {
                return;
            }

            if (EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.PowerInfo] == null)
            {
                return;
            }

            List<string> output_clock_pins = new List<string>();
            for (int i = 2; i <= nWireWorksheet.Dimension.End.Column; i++)
            {
                output_clock_pins.Add(EpplusExtensions.GetCellValue(nWireWorksheet, 1, i));
            }


            PowerInfoSheet powerInfo = TestPlanStatic.PowerInfoSheet;
            foreach (PowerInfoRow row in powerInfo.Rows)
            {
                if (output_clock_pins.Contains(row.PinName))
                {
                    var paPinn = new ProtocolAwarePin();
                    if (row.PinName.Contains("::"))
                    {
                        //Differential pin. eg. XI0::XO0
                        paPinn.OutClk = _regex4.Match(row.PinName).Groups[CommonConst.OutFirst].ToString();
                        paPinn.OutClkDiff = _regex4.Match(row.PinName).Groups[CommonConst.OutSecond].ToString();
                    }
                    else
                    {
                        //single-end pin. eg. RT_CLK32768
                        paPinn.OutClk = row.PinName;
                    }

                    ProtocolAwarePin target = SettingInfo.NwirePins.FirstOrDefault(p => p.OutClk == paPinn.OutClk);
                    if (target != null)
                    {
                        target.OutClk = paPinn.OutClk;
                        target.PowerUpSeq = row.PowerSequence;
                        if (!string.IsNullOrEmpty(row.PowerDownSequence))
                        {
                            target.PowerDownSeq = row.PowerDownSequence;
                        }
                    }
                }
            }

        }

        public Dictionary<string, string> GetPatternDic()
        {
            var dic = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int cnt = 1;
            for (int i = 0; i < SettingInfo.SettingTable.Rows.Count; i++)
            {
                string patternName = SettingInfo.SettingTable.Rows[i][0].ToString().ToUpper();
                TryAddPattern(dic, patternName, ref cnt);
            }
            return dic;
        }
        internal static bool TryAddPattern(
            Dictionary<string, string> dic,
            string patternName,
            ref int cnt)
        {
            if (patternName.Length <= 20)
            {
                return false;
            }

            if (dic.ContainsKey(patternName))
            {
                return false;
            }


            const string shortName = "Pattern";
            dic.Add(patternName, shortName + "_" + cnt);
            cnt++;
            return true;
        }

        private void SetIoVoltage()
        {
            if (LocalSpecs.TestPlanFileName.Equals("N/A"))
            {
                return;
            }
            ExcelWorksheet ioContiSheet = EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.ContiIo];
            ExcelWorksheet ioLevelSheet = EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.IoLevels];
            IoContiSheet contiSheet = null;
            if (ioContiSheet != null)
            {
                var contiReader = new IoContiReader();
                contiSheet = contiReader.ReadSheet(ioContiSheet);
            }
            else if (ioLevelSheet != null)
            {
                var ioLevelsSheetReader = new IoLevelsSheetReader();
                IoLevelsSheet ioLevels = ioLevelsSheetReader.ReadSheet(ioLevelSheet);
                contiSheet = ioLevels.ConvertIoContiSheet();
            }

            if (contiSheet == null)
            {
                return;
            }

            foreach (EnumEquipment equipment in TestPlanStatic.Equipments)
            {
                foreach (ProtocolAwarePin nWirePin in SettingInfo.NwirePins)
                {
                    string ioVoltage = "";
                    if (contiSheet.IsPinExist(nWirePin.OutClk))
                    {
                        ioVoltage = contiSheet.GetVoltage(nWirePin.OutClk);
                    }
                    else if (contiSheet.IsPinExist(nWirePin.CreatePaClkPinName(equipment)))
                    {
                        ioVoltage = contiSheet.GetVoltage(nWirePin.CreatePaClkPinName(equipment));
                    }
                    if (ioVoltage == "")
                    {
                        continue;
                    }

                    string targetValue = ResolveVoltage(ioVoltage);
                    nWirePin.OutClkVoltage = GetOutClkVoltage(targetValue, nWirePin.OutClkVoltage);
                }
            }
        }

        internal static string ResolveVoltage(string ioVoltage)
        {
            string value = _regex5.Match(ioVoltage).Groups["value"].ToString();
            string unit = _regex5.Match(ioVoltage).Groups["unit"].ToString();

            string targetValue;
            if (unit.Equals(""))
            {
                targetValue = value;
            }
            else
            {
                value.TryCombineVolt(unit, out targetValue);
            }

            return targetValue;
        }

        internal static double GetOutClkVoltage(string targetValue, double currentValue)
        {
            double voltage;
            if (!double.TryParse(targetValue, out voltage))
            {
                return currentValue;
            }
            return currentValue == 0d ? voltage : currentValue;
        }


        public void SetSuperClock(Dictionary<string, ChannelMapSheet> channelMapSheets)
        {
            string path = Path.Combine(LocalSpecs.SettingFolder, "Config", "CurrentChannelMap.txt");
            if (!File.Exists(path))
            {
                return;
            }

            if (channelMapSheets == null || channelMapSheets.Count == 0)
            {
                return;
            }

            var currentChannelMap = new CurrentChannelReader();
            CurrentChannelSheet currentChannelSheet = CurrentChannelReader.ReadFile(path);
            Dictionary<string, string> dictionary = currentChannelSheet.GetPogoMapping("HSDU, HSD-U");
            var list = new List<List<string>>();
            foreach (KeyValuePair<string, ChannelMapSheet> sheet in channelMapSheets)
            {
                if (sheet.Value.Rows.Any(x => x.Type.Equals("I/O", StringComparison.OrdinalIgnoreCase)))
                {
                    var channelMapRows = sheet.Value.Rows.Where(x => x.Type.Equals("I/O", StringComparison.OrdinalIgnoreCase)).ToList();
                    if (sheet.Value.IsPogo)
                    {
                        channelMapRows = ConvertPogo2Single(channelMapRows, dictionary);
                    }

                    list.Add(GenSuperClkPin(channelMapRows));
                }
            }

            //Get common for all sheet
            if (list.Count == 0)
            {
                return;
            }

            var pins = new List<string>();
            foreach (string item in list[0])
            {
                bool flag = list.All(sheet => sheet.Any(x => x.Equals(item, StringComparison.OrdinalIgnoreCase)));
                if (flag)
                {
                    pins.Add(item);
                }
            }
        }

        internal List<ChannelMapRow> ConvertPogo2Single(List<ChannelMapRow> channelMapRows, Dictionary<string, string> pogoMapping)
        {
            foreach (ChannelMapRow row in channelMapRows)
            {
                for (int i = 0; i < row.Sites.Count; i++)
                {
                    string data = row.Sites[i];
                    if (data.Contains('.'))
                    {
                        string[] array = data.Split('.');
                        if (array.Length == 2)
                        {
                            string ch = array[1].ToLower();
                            if (pogoMapping.TryGetValue(ch, out string value))
                            {
                                row.Sites[i] = data.Split('.')[0] + "." + value;
                            }
                        }
                    }
                }
            }
            return channelMapRows;
        }

        private List<string> GenSuperClkPin(List<ChannelMapRow> channelMapRows)
        {

            var list = new List<string>();
            foreach (ProtocolAwarePin nwire in SettingInfo.NwirePins)
            {
                if (nwire.Freq > ProtocolAwarePin.MaxFreq)
                {
                    list.Add(nwire.OutClk);
                    continue;
                }

                //Check frequency (<250MHz => normal clock,250MHz ~ 550MHz Need to confirm)
                if (nwire.Freq < ProtocolAwarePin.MinFreqSuperClock)
                {
                    continue;
                }

                string superClk = nwire.OutClk + "_" + ProtocolAwarePin.ConPa;
                string refClk = ProtocolAwarePin.ConRefClk + "_" + nwire.OutClk;
                bool isSuperClkPin = true;
                if (!channelMapRows.Any(x => x.DeviceUnderTestPinName.Equals(superClk, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (!channelMapRows.Any(x => x.DeviceUnderTestPinName.Equals(refClk, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                ChannelMapRow superClkRow = channelMapRows.First(x => x.DeviceUnderTestPinName.Equals(superClk, StringComparison.OrdinalIgnoreCase));

                #region Check superClk must be even ch & odd ch is empty
                for (int i = 0; i < superClkRow.Sites.Count; i++)
                {
                    string slot = "";
                    int ch = 0;
                    if (ParseChannel(superClkRow.Sites[i], ref slot, ref ch))
                    {
                        string oddCh = slot + ".ch" + (ch - 1);
                        if (!(ch % 2 == 0 && !channelMapRows.SelectMany(y => y.Sites).Any(x => x.Equals(oddCh))))
                        {
                            isSuperClkPin = false;
                        }
                    }
                }
                #endregion

                #region Check superClk and refClk channel block
                if (isSuperClkPin)
                {
                    //check super clk and ref_clk row
                    var superClkSite = channelMapRows.Where(x => x.DeviceUnderTestPinName.Equals(superClkRow.DeviceUnderTestPinName)
                                                                 || x.DeviceUnderTestPinName.StartsWith(refClk, StringComparison.OrdinalIgnoreCase)).SelectMany(y => y.Sites)
                        .Where(z => _regex7.IsMatch(z)).ToList();

                    if (superClkSite.Any(x => !IsInTheSameChannelBlock(superClkRow.Sites[0], x)))
                    {
                        isSuperClkPin = false;
                    }

                    //check other row
                    if (isSuperClkPin)
                    {
                        var otherSite = channelMapRows.Where(x => !(x.DeviceUnderTestPinName.Equals(superClkRow.DeviceUnderTestPinName)
                                                                    || x.DeviceUnderTestPinName.StartsWith(refClk, StringComparison.OrdinalIgnoreCase))).SelectMany(y => y.Sites)
                            .Where(z => _regex7.IsMatch(z)).ToList();

                        if (otherSite.Any(t => IsInTheSameChannelBlock(superClkRow.Sites[0], t)))
                        {
                            isSuperClkPin = false;
                        }
                    }
                }
                #endregion

                if (isSuperClkPin)
                {
                    list.Add(nwire.OutClk);
                }
            }
            return list;
        }

        internal bool ParseChannel(string data, ref string slot, ref int ch)
        {
            bool flag = false;
            string[] array = data.Split('.');
            if (array.Length == 2)
            {
                slot = array[0];
                ch = int.Parse(data.Split('.')[1].Replace("ch", ""));
                flag = true;
            }
            return flag;
        }

        private bool IsInTheSameChannelBlock(string data1, string data2)
        {
            string slot1 = "";
            int ch1 = 0;
            string slot2 = "";
            int ch2 = 0;
            ParseChannel(data1, ref slot1, ref ch1);
            ParseChannel(data2, ref slot2, ref ch2);

            return IsInTheSameChannelBlockCore(slot1, ch1, slot2, ch2);
        }

        internal bool IsInTheSameChannelBlockCore(
            string slot1, int ch1, string slot2, int ch2)
        {
            return slot1 == slot2 && ch1 / 32 == ch2 / 32;
        }

        public List<string> ReferenceFlow()
        {
            var names = new List<string>();
            DataTable table = Instance().SettingInfo.SettingTable;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                string item = table.Rows[i][0].ToString();
                if (ShouldAddReferenceFlow(item))
                {
                    names.Add(item);
                }

            }
            return names;
        }

        internal bool ShouldAddReferenceFlow(string item)
        {
            return !_regex.IsMatch(item);
        }

    }
}
