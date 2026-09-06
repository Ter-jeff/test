using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using OfficeOpenXml;

using TestPlanLib.Const;

namespace Automation.Reader
{
    public class NwireReader
    {
        #region Field
        private ExcelWorksheet _worksheet;
        private int _startRow = 1;
        private int _startColumn = 1;

        private int _pinRow = -1;               //Output Clock Pin
        private int _frequencyRow = -1;
        private int _relayControlRow = -1;
        private int _refClkPinRow = -1;
        private int _refClkVoltage = -1;
        private int _outClkPinVoltage = -1;
        private int _protocolRow = -1;
        private int _dcCategoryRow = -1;

        private readonly List<int> _setRowList = new List<int>();
        private NwireSetting _settingInfo = new NwireSetting();
        #endregion

        #region Member Function

        public NwireSetting ReadFlow(ExcelWorksheet nWireWorksheet)
        {
            try
            {
                _worksheet = nWireWorksheet;

                ReadHeader();

                ReadData();

            }
            catch (Exception e)
            {

                throw new Exception("Error occurs during Reading nWire setting file: " + e.StackTrace);

            }
            return _settingInfo;
        }

        private void ReadHeader()
        {
            string header;
            bool hasFind = false;
            for (int i = 1; i <= _worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= _worksheet.Dimension.End.Column; j++)
                {
                    header = EpplusExtensions.GetCellValue(_worksheet, i, j);
                    if (Regex.IsMatch(header, TestPlanConst.OutClkPinPattern, RegexOptions.IgnoreCase))
                    {
                        hasFind = true;
                        _startRow = i;
                        _startColumn = j;
                        break;
                    }
                }
                if (hasFind)
                {
                    break;
                }
            }
            _pinRow = _startRow;

            for (int i = _startRow + 1; i <= _worksheet.Dimension.End.Row; i++)
            {
                header = EpplusExtensions.GetCellValue(_worksheet, i, _startColumn);
                if (header.Equals(""))
                {
                    break;
                }

                if (Regex.IsMatch(header, TestPlanConst.FrequencyPattern, RegexOptions.IgnoreCase))
                {
                    _frequencyRow = i;
                }
                else if (Regex.IsMatch(header, TestPlanConst.RelayControlPattern, RegexOptions.IgnoreCase))
                {
                    _relayControlRow = i;
                }
                else if (Regex.IsMatch(header, TestPlanConst.RefClkPinPattern, RegexOptions.IgnoreCase))
                {
                    _refClkPinRow = i;
                }
                else if (Regex.IsMatch(header, TestPlanConst.OutClkPinVoltagePattern, RegexOptions.IgnoreCase))
                {
                    _outClkPinVoltage = i;
                }
                else if (Regex.IsMatch(header, TestPlanConst.RefClkVoltagePattern, RegexOptions.IgnoreCase))
                {
                    _refClkVoltage = i;
                }
                else if (Regex.IsMatch(header, TestPlanConst.Protocol, RegexOptions.IgnoreCase))
                {
                    _protocolRow = i;
                }
                else if (Regex.IsMatch(header, TestPlanConst.DcCategory, RegexOptions.IgnoreCase))
                {
                    _dcCategoryRow = i;
                }
                else
                {
                    _setRowList.Add(i);
                }
            }
        }

        private void ReadData()
        {
            _settingInfo = new NwireSetting();
            DataTable settingTable = new DataTable();
            settingTable.Columns.Add("Pin");
            if (HasMissingRequiredSetting(_pinRow, _frequencyRow))
            {
                throw new Exception("The Setting sheet missing required information!");
            }
            for (int i = _startColumn + 1; i <= _worksheet.Dimension.End.Column; i++)
            {
                ProtocolAwarePin paPin = new ProtocolAwarePin();

                string pin = EpplusExtensions.GetCellValue(_worksheet, _pinRow, i);
                if (pin.Equals(""))
                {
                    break;
                }

                if (pin.Contains("::"))
                {
                    //Differential pin. eg. XI0::XO0
                    paPin.OutClk = Regex.Match(pin, TestPlanConst.PinRegPattern).Groups[TestPlanConst.OutFirst].ToString();
                    paPin.OutClkDiff = Regex.Match(pin, TestPlanConst.PinRegPattern).Groups[TestPlanConst.OutSecond].ToString();
                    paPin.PinType = IoPinType.Diff;
                }
                else
                {
                    //single-end pin. eg. RT_CLK32768
                    paPin.OutClk = pin;
                    paPin.PinType = IoPinType.Single;
                }

                string frequency = EpplusExtensions.GetCellValue(_worksheet, _frequencyRow, i);
                string value = Regex.Match(frequency, TestPlanConst.UnitRegPattern).Groups[TestPlanConst.Value].ToString();
                string unit = Regex.Match(frequency, TestPlanConst.UnitRegPattern).Groups[TestPlanConst.Unit].ToString();
                if (value.TryCombineHz(unit, out string targetFreq))
                {
                    paPin.Freq = double.Parse(targetFreq);
                }

                if (_relayControlRow == -1)
                {
                    //Add default relay name
                    paPin.RelayControl = "Relay_" + paPin.OutClk;
                }
                else
                {
                    //Read Relay from config sheet
                    paPin.RelayControl = EpplusExtensions.GetCellValue(_worksheet, _relayControlRow, i);
                }

                if (_refClkPinRow != -1)
                {
                    string refClk = EpplusExtensions.GetCellValue(_worksheet, _refClkPinRow, i);
                    paPin.RefClk = string.IsNullOrEmpty(refClk) ? "" : refClk;
                }

                if (_outClkPinVoltage != -1)
                {
                    string targetValue = EpplusExtensions.GetCellValue(_worksheet, _outClkPinVoltage, i);
                    ResolveVoltage(targetValue, v => paPin.OutClkVoltage = v);
                }

                if (_refClkVoltage != -1)
                {
                    string targetValue = EpplusExtensions.GetCellValue(_worksheet, _refClkVoltage, i);
                    ResolveVoltage(targetValue, v => paPin.RefClkVoltage = v);
                }

                if (_protocolRow != -1)
                {
                    paPin.Protocol = EpplusExtensions.GetCellValue(_worksheet, _protocolRow, i);
                }

                if (_dcCategoryRow != -1)
                {
                    paPin.DcCategory = EpplusExtensions.GetCellValue(_worksheet, _dcCategoryRow, i);
                }

                settingTable.Columns.Add(paPin.OutClk);
                _settingInfo.NwirePins.Add(paPin);
            }

            //To add new row to specify Flow_nwire_Default or Flow_nwire_HARDIP
            foreach (int i in _setRowList)
            {
                DataRow dataRow = settingTable.NewRow();
                string type = EpplusExtensions.GetCellValue(_worksheet, i, _startColumn);
                if (!string.IsNullOrEmpty(type))
                {
                    type = Regex.Replace(type, "FlowControl:", "", RegexOptions.IgnoreCase);
                }

                dataRow[0] = Regex.Replace(type, $"^{NwireSetting.ConFlownWire.TrimEnd('_')}_?", "");
                int count = 1;
                for (int j = _startColumn + 1; j <= _worksheet.Dimension.End.Column; j++)
                {
                    if (EpplusExtensions.GetCellValue(_worksheet, _startRow, j).Equals(""))
                    {
                        break;
                    }
                    string flag = EpplusExtensions.GetCellValue(_worksheet, i, j);
                    dataRow[count++] = flag;
                }
                settingTable.Rows.Add(dataRow);
            }

            string flagReset = "Reset";
            object[] values = new object[_worksheet.Dimension.End.Column];
            for (int i = 1; i <= _settingInfo.NwirePins.Count; i++)
            {
                values[i] = flagReset;
            }
            settingTable.Rows.Add(values);

            _settingInfo.SettingTable = settingTable;
            _settingInfo.PatternDic = GetPatternDic();
        }

        internal static bool HasMissingRequiredSetting(
            int pinRow,
            int frequencyRow)
        {
            return pinRow == -1
                || frequencyRow == -1;
        }

        internal static void ResolveVoltage(
            string targetValue,
            Action<double> setVoltage)
        {
            if (double.TryParse(targetValue, out double voltage))
            {
                setVoltage(voltage);
            }
        }


        public Dictionary<string, string> GetPatternDic()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int cnt = 1;
            for (int i = 0; i < _settingInfo.SettingTable.Rows.Count; i++)
            {
                string patternName = _settingInfo.SettingTable.Rows[i][0].ToString().ToUpper();
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

        #endregion

    }

    /// <summary>
    /// All the information from nWire setting file
    /// 
    /// Datatable:
    /// Pin              XI0              RT_CLK32768
    /// Default         Enable             Enable
    /// HardIP          Disable            Disable                  
    /// </summary>
    public class NwireSetting
    {
        #region Field

        private static readonly Regex _regex = new Regex("Default", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        #endregion

        #region Properity
        public const string RelayEnableNwire = "Relay_Enable_Nwire";
        public const string RelayDisableNwire = "Relay_Disable_Nwire";
        public const string ConFlownWire = "Flow_nWire_";

        public List<string> ReferenceFlow(DataTable table)
        {
            var names = new List<string>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                string item = table.Rows[i][0].ToString();
                if (!_regex.IsMatch(item))
                {
                    names.Add(item);
                }
            }
            return names;
        }

        public List<ProtocolAwarePin> NwirePins { set; get; } = new List<ProtocolAwarePin>();

        public DataTable SettingTable { set; get; }

        public Dictionary<string, string> PatternDic { set; get; }

        public double SupportBoardFreq { get; set; }

        #endregion

        #region Constructor
        public NwireSetting()
        {
            NwirePins = new List<ProtocolAwarePin>();
            SettingTable = new DataTable();
            PatternDic = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
        #endregion

        #region Member Function

        public FlowRow GetNwirePatternCall(string patternCall)
        {
            var row = new FlowRow { Opcode = "Call", Parameter = ConFlownWire + patternCall };
            return row;
        }

        public FlowRow GetNwireCall(string flowSheetName)
        {
            string opcode = "Call";
            var row = new FlowRow { Opcode = opcode };
            string keyword;
            bool find = false;
            for (int i = 0; i < SettingTable.Rows.Count; i++)
            {
                keyword = SettingTable.Rows[i][0].ToString();
                if (Regex.IsMatch(flowSheetName, keyword, RegexOptions.IgnoreCase))
                {
                    row.Parameter = ConFlownWire + keyword;
                    find = true;
                }
                if (!find)
                {
                    row.Parameter = "Flow_nWire_Default";
                }
            }
            return row;
        }

        #endregion
    }

    /// <summary>
    /// PA Engin out clock pin and reference clock pin and some properities
    /// </summary>
    public class ProtocolAwarePin
    {
        #region Field
        public const string ConPa = "PA";
        public const string ConRefClk = "REFCLK";
        public const string ConPort = "Port";
        public const string ConDiff = "Diff";
        public const string ConGlb = "GLB";
        public const string ConPins = "Pins";
        public const string ConFreq = "Freq";
        public const string ConPowerSequence = "PowerSequence";
        public const string ConPowerDownSequence = "PowerDownSequence";
        public const string ConVar = "VAR";
        /* The digital Channel has some limit for Frequence, if the Freq exceed the limit, the tool will not change Time Set*/
        public const double MaxFreq = 550e6;             //550MHZ             //Digital Channel Maximun Frequence
        public const double MinFreq = 700e3;             //700KHZ             //Digital Channel Minimun Frequence
        public const double MinFreqSuperClock = 250e6;   //550MHZ             //Super Clock Minimun Frequence
        public const string ConValidFamilies = "validFamilies";
        public const string CondefaultName = "defaultName";
        public const string ConName = "name";
        #endregion

        #region Properity
        //Reference Clock pin. Example: REFCLK_XI0
        public string RefClk { set; get; } = string.Empty;

        //Only when The pin is Diff, this value is valid. Example: XO0
        public string OutClkDiff { set; get; } = string.Empty;

        //Out Clock pin. Example: XI0
        public string OutClk { set; get; } = string.Empty;

        //Differential or Single-end
        public IoPinType PinType { set; get; } = IoPinType.Single;

        //Target frequency of output clock pin
        public double Freq { set; get; }

        //Power sequence of nWire pin
        public string PowerUpSeq { set; get; } = "99";

        public string PowerDownSeq { set; get; }
        //Voltage of output Clock pin
        public double OutClkVoltage { set; get; }

        //Voltage of reference  Clock pin
        public double RefClkVoltage { set; get; } = 1.8; //This value is depend on Support Board, it can be always 1.8v.

        //Relay to control PA pin and Digital channgel, Relay_XI0
        public string RelayControl { set; get; } = string.Empty;

        //ExtraPin for LCD
        public string ExtraPin { set; get; }

        public string FlowControlAction { set; get; }

        public string ControlAction { set; get; }

        public string Protocol { set; get; }

        public string DcCategory { set; get; }
        #endregion

        #region Member Function

        public string CreatePinNameWithDiff()
        {
            if (PinType == IoPinType.Diff)
            {
                return OutClk + "_" + ConDiff;
            }
            return OutClk;
        }

        public string CreatePaClkPinName(EnumEquipment equipment)
        {
            if (equipment == EnumEquipment.UltraFlexPlus)
            {
                return OutClk;
            }

            return OutClk + "_" + ConPa;
        }

        public string CreateFamily(XmlNode node)
        {
            if (node.Attributes != null)
            {
                return node.Attributes[ConValidFamilies].Value;
            }

            return "";
        }

        public string CreateType(XmlNode node)
        {
            if (node.Attributes != null)
            {
                return node.Attributes[CondefaultName].Value;
            }

            return "";
        }

        /// <summary>
        /// Create the real PA pin name eg. "XO0_PA"
        /// </summary>
        /// <returns></returns>
        public string CreatePaClkDiffPinName(EnumEquipment equipment)
        {
            if (equipment == EnumEquipment.UltraFlexPlus)
            {
                return OutClkDiff;
            }

            return OutClkDiff + "_" + ConPa;
        }

        /// <summary>
        /// Create Port Name for PA Pin eg. "XI0_Diff_Port", "RT_CLK32768_Port"
        /// </summary>
        /// <returns></returns>
        public string CreatePortName(EnumEquipment equipment)
        {
            if (PinType.Equals(IoPinType.Diff))
            {
                if (equipment == EnumEquipment.UltraFlexPlus)
                {
                    return OutClk + "_" + ConDiff;
                }

                return OutClk + "_" + ConDiff + "_" + ConPort;
            }
            if (equipment == EnumEquipment.UltraFlexPlus)
            {
                return OutClk;
            }

            return OutClk + "_" + ConPort;
        }

        /// <summary>
        /// Create differential pin group name for Differential pin
        /// eg. "XI0_Diff_PA"
        /// </summary>
        /// <returns></returns>
        public string CreateDiffPinGroupName(EnumEquipment equipment)
        {
            if (PinType.Equals(IoPinType.Diff))
            {
                if (equipment == EnumEquipment.UltraFlexPlus)
                {
                    return OutClk + "_" + ConDiff;
                }

                return OutClk + "_" + ConDiff + "_" + ConPa;
            }
            return "";
        }

        /// <summary>
        /// Create Global spec symbol for frequency value
        /// eg. "XI0_Diff_Freq_GLB", "RT_CLK32768_Freq_GLB"
        /// </summary>
        /// <returns></returns>
        public string CreateFreqSpecName(EnumEquipment equipment)
        {
            return CreatePortName(equipment) + "_" + ConFreq + "_" + ConGlb;
        }

        /// <summary>
        /// Create AC spec variable for frequency value
        /// eg. "XI0_Diff_Freq_GLB", "RT_CLK32768_Freq_VAR"
        /// </summary>
        /// <returns></returns>
        public string CreateFreqVarName()
        {
            if (PinType.Equals(IoPinType.Diff))
            {
                return OutClk + "_" + ConDiff + "_" + ConFreq + "_" + ConVar;
            }
            return OutClk + "_" + ConFreq + "_" + ConVar;
        }

        /// <summary>
        /// Create Io voltage Global spec symbol for output clock pin 
        /// </summary>
        /// <returns></returns>
        public string CreateOutClkLevelSpecName()
        {
            if (PinType.Equals(IoPinType.Diff))
            {
                return ConPins + "_" + OutClk + "_" + ConDiff + "_" + ConGlb;
            }
            return ConPins + "_" + OutClk + "_" + ConGlb;
        }

        /// <summary>
        /// Create Io voltage Global spec symbol for reference clock pin 
        /// </summary>
        /// <returns></returns>
        public string CreateRefClkLevelSpecName()
        {
            if (PinType.Equals(IoPinType.Diff))
            {
                return ConPins + "_" + RefClk + "_" + ConDiff + "_" + ConGlb;
            }
            return ConPins + "_" + RefClk + "_" + ConGlb;
        }

        /// <summary>
        /// Create Power sequence for the nWire pin
        /// eg. "XI0_Diff_Port_PowerSequence_GLB", "RT_CLK32768_Port_PowerSequence_GLB"
        /// </summary>
        /// <returns></returns>
        public string CreatePowerSeqSpecName(EnumEquipment equipment)
        {
            return CreatePortName(equipment) + "_" + ConPowerSequence + "_" + ConGlb;
        }

        public string CreatePowerDownSeqSpecName(EnumEquipment equipment)
        {
            return CreatePortName(equipment) + "_" + ConPowerDownSequence + "_" + ConGlb;
        }

        public string CreateAcSymbol()
        {
            if (PinType == IoPinType.Diff)
            {
                return OutClk + "_" + ConDiff + "_" + ConFreq;
            }
            return OutClk + "_" + ConFreq;
        }
        #endregion
    }

    /// <summary>
    /// Differential or Single-end
    /// </summary>
    public enum IoPinType
    {
        Diff,
        Single
    }
}
