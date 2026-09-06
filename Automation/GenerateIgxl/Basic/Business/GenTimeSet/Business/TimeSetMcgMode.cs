using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

using Automation.Reader;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using TestPlanLib.Basic;

namespace Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business
{
    public class VpnSettingFile
    {
        public string Extensions;
        public string ConfigFile;
        public string DicOrgStr;
        public string DicSpecStr;
    }

    /// <summary>
    /// Set time set as MCG Mode:
    ///     1. Set period of Pin XI0, RT_CLK32768 as the same value with XI0_PA, RT_CLK32768_PA;
    ///     2. Set Clock Set Up as "Clock", Data source as "AllHi" and format as "RL";
    ///     3. Set Drive data as Period/2 and Data Return as Period;
    ///     4. Check if there exist Differential pins for differential nWire pin, such as XI0 and XO0;
    ///     5. Differential XO0 set Data Source with "AllLo" and format as "RH".
    /// </summary>
    public class TimeSetMcgMode
    {
        #region Field

        private const string SetUpClock = "clock";
        private const string SetUpClock2X = "clock_2X";
        private const string SrcAllHi = "ALLHI";
        private const string SrcAllLo = "ALLLO";
        private const string FmtRl = "RL";
        private const string FmtRh = "RH";
        private const string Off = "Off";
        private readonly List<ProtocolAwarePin> _nWirePins;          //nWire pin List
        private VpnSettingFile _settingFile;
        private Dictionary<string, string> _patternDigi2TypeDic;
        private Dictionary<string, string> _patternDigi4TypeDic;

        private const string HardIp = "HARD_IP";

        /* The digital Channel has some limit for Frequence, if the 
         Freq exceed the limit, the tool will not change Time Set*/
        private readonly double _digitalChanneMaxFreq = 550e6;   //550MHZ             //Digital Channel Maximun Frequence
        private readonly double _digitalChannelMinFreq = 700e3;  //700KHZ             //Digital Channel Minimun Frequence
        private static readonly Regex _regex = new Regex(@"(?<key>\w+)\s*[:]\s*(?<value>\w+)");
        private static readonly Regex _regex2 = new Regex(@"(?<key>\w+)\s*[:]\s*(?<value>\w+)");
        #endregion

        #region Constructor

        public TimeSetMcgMode(List<ProtocolAwarePin> nWirePins)
        {
            _nWirePins = nWirePins;

            ReadCfg();

        }
        #endregion

        public void ConverFlow(List<ComTimeSetBasicSheet> timeSetSheets, List<PatternData> patterns)
        {

            foreach (TimeSetBasicSheet basicSheet in timeSetSheets)
            {
                PatternData row = patterns.Find(p => p.TimeSetVersion.Equals(basicSheet.Name));
                if (row != null && !JudgeHardIpPattern(row.PatternName))
                {
                    continue;
                }
                foreach (TSet tset in basicSheet.Rows)
                {
                    foreach (ProtocolAwarePin wirePin in _nWirePins)
                    {
                        if (tset.TimingRows.Exists(p => p.PinGrpName.Equals(wirePin.OutClk, StringComparison.OrdinalIgnoreCase)))
                        {
                            ModifyMcgMode(tset, wirePin);
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Read the Config of VPN from Config file
        /// </summary>
        private void ReadCfg()
        {
            string infoConfig = Path.Combine(AppContext.BaseDirectory, "Config", "VpnSettingFile.xml");
            var reader = new XmlSerializer(typeof(VpnSettingFile));
            var cfgXml = new StreamReader(infoConfig);
            _settingFile = (VpnSettingFile)reader.Deserialize(cfgXml);
            _settingFile.DicOrgStr =
                _settingFile.DicOrgStr.Replace("\t", "").Replace("\n", "").Replace("\r", "").Replace(" ", "").Replace("\'", "").Replace("\"", "");
            _settingFile.DicSpecStr =
                _settingFile.DicSpecStr.Replace("\t", "").Replace("\n", "").Replace("\r", "").Replace(" ", "").Replace("\'", "").Replace("\"", "");
            cfgXml.Close();

            _patternDigi2TypeDic = new Dictionary<string, string>();
            _patternDigi4TypeDic = new Dictionary<string, string>();

            List<string> digi2List = _settingFile.DicOrgStr.Split(',').ToList();
            List<string> digi4List = _settingFile.DicSpecStr.Split(',').ToList();
            foreach (string s in digi2List)
            {
                string key = _regex.Match(s).Groups["key"].ToString();
                string value = _regex.Match(s).Groups["value"].ToString();
                _patternDigi2TypeDic.Add(key, value);
            }

            foreach (string s in digi4List)
            {
                string key = _regex2.Match(s).Groups["key"].ToString();
                string value = _regex.Match(s).Groups["value"].ToString();
                _patternDigi4TypeDic.Add(key, value);
            }
        }

        private bool JudgeHardIpPattern(string pattern)
        {
            List<string> subName = pattern.Split('_').ToList();
            if (subName.Count > 5)
            {
                if (_patternDigi2TypeDic.ContainsKey(subName[2].ToUpper()) &&
                    _patternDigi2TypeDic[subName[2].ToUpper()] == HardIp)
                {
                    return true;
                }

                if (_patternDigi4TypeDic.ContainsKey(subName[4].ToUpper()) &&
                    _patternDigi4TypeDic[subName[4].ToUpper()] == HardIp)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// if nWire pin is differential pin, Update two pins: XI0, XO0
        /// else only update one pin: RT_CLK32768
        /// </summary>
        /// <param name="tset">Tset</param>
        /// <param name="nWirePin">nWirePin</param>
        private void ModifyMcgMode(TSet tset, ProtocolAwarePin nWirePin)
        {
            string period = "= 1/_" + nWirePin.CreateFreqVarName();//CreateFreqSpecName();
            string driveData = "= 1/(2*_" + nWirePin.CreateFreqVarName() + ")";//CreateFreqSpecName();

            if (nWirePin.Freq < _digitalChannelMinFreq)
            {
                //If the frequece exceed the Low limit, do not need to change time set
                return;
            }

            TimingRow timingRow = tset.TimingRows.Find(p => p.PinGrpName.Equals(nWirePin.OutClk, StringComparison.OrdinalIgnoreCase));
            timingRow.PinGrpSetup = SetUpClock;

            if (nWirePin.Freq > _digitalChanneMaxFreq)
            {
                //If the frequence exceed the high limit, change the Mode as "clock_2X"
                period = "= 2/_" + nWirePin.CreateFreqVarName();//CreateFreqSpecName();
                driveData = "= 1/_" + nWirePin.CreateFreqVarName();//CreateFreqSpecName();
                timingRow.PinGrpSetup = SetUpClock2X;
            }

            /*
             t0t1       t2           t3t0t1     t2          t3
                          ____________           ____________ 
             _____________            ___________
             */
            timingRow.PinGrpClockPeriod = period;
            timingRow.DataSrc = SrcAllHi;
            timingRow.DataFmt = FmtRl;
            timingRow.DriveData = driveData;
            timingRow.DriveReturn = period;
            timingRow.CompareMode = Off;
            string outClkEdgeMode = timingRow.EdgeMode;

            if (nWirePin.PinType == IoPinType.Diff)
            {
                timingRow =
                    tset.TimingRows.Find(
                        p => p.PinGrpName.Equals(nWirePin.OutClkDiff, StringComparison.OrdinalIgnoreCase));
                if (timingRow == null)
                {
                    timingRow = new TimingRow { PinGrpName = nWirePin.OutClkDiff };
                    tset.AddTimingRow(timingRow);
                }
                /*
             t0t1       t2           t3t0t1     t2          t3
             _____________            ____________            
                           ___________            ____________
             */
                timingRow.PinGrpSetup = SetUpClock;
                timingRow.PinGrpClockPeriod = period;
                timingRow.DataSrc = SrcAllLo;
                timingRow.DataFmt = FmtRh;
                timingRow.DriveData = driveData;
                timingRow.DriveReturn = period;
                timingRow.CompareMode = Off;
                timingRow.EdgeMode = outClkEdgeMode;
            }
        }

    }
}
