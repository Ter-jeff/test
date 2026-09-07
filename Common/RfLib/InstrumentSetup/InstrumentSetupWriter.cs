using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.Extension;

using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using RfLib.InstrumentSetup.InstrumentTypeData;

using DataColumn = System.Data.DataColumn;
using DataTable = System.Data.DataTable;

namespace RfLib.InstrumentSetup
{

    public class InstrumentSetupWriter
    {
        private const string ConHeaderRfinstrumentsetup = "RF Instrument Setup";
        private const string ConHeaderSubsetting = "SubSetting";
        private const string ConHeaderMeastype = "MeasType";
        private const string ConHeaderInstrumenttype = "Instrument Type";

        private const int DefaultTypeIndex = 0;
        private const int DefaultHeaderIndex = 1;
        private const int DefaultValueIndex = 2;
        private const int EndInstrumentTypeSettingNumber = 27;

        private int _startColNumber = -1;
        private int _endColNumber = -1;
        private readonly HeaderColumnIndices _headerIndices = new();

        private ExcelWorksheet _excelWorksheet = null!;
        private readonly InstrumentTypeSelector _typeSelector;
        private InstrumentConfigSheet _instrumentSheet = new("");

        public Dictionary<string, int> TypeIndexList { get; set; } = [];
        public SortedList<int, string> DefaultValue { get; set; } = [];
        public List<InstrumentTypeSetting> InstTypeSetting { get; set; } = [];
        public List<InstrumentTypePara> TypeParameterHeaderList { get; set; } = [];
        public List<ReplaceInstrumentClass> ReplaceInstKey { get; set; } = [];

        public InstrumentSetupWriter(List<ChannelMapSheet> channelMapSheets)
        {
            _typeSelector = new InstrumentTypeSelector(channelMapSheets);
        }

        public InstrumentSetupWriter()
        {
            _typeSelector = new InstrumentTypeSelector([]);
        }

        public void ReadConfigSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return;
            }

            _excelWorksheet = excelWorksheet;

            _instrumentSheet = new InstrumentConfigSheet(excelWorksheet.Name);

            GetDimensions();
            GetHeaderIndex();

            TypeIndexList = InstrumentTypeUtility.GetInstrumentType(_excelWorksheet, _endColNumber);
            TypeParameterHeaderList = InstrumentTypeUtility.GetInstrumentParameterHeader(_excelWorksheet, TypeIndexList, _endColNumber);
            DefaultValue = ReadDefaultValue();
            InstTypeSetting = ReadInstrumentTypeSetting();
        }

        public void ReadReplaceConfigSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                ReplaceInstKey = [];
                return;
            }

            ReplaceInstKey = new ReplaceInstrumentReader().ReadSheet(excelWorksheet)?.ReplaceClass ?? [];
        }

        public DataTable WriteDefaultHeader()
        {
            var dt = new DataTable();
            for (int i = 1; i <= _endColNumber; i++)
            {
                dt.Columns.Add(new DataColumn());
            }
            for (int i = 1; i <= 4; i++)
            {
                dt.Rows.Add(dt.NewRow());
            }

            foreach (KeyValuePair<string, int> type in TypeIndexList)
            {
                dt.Rows[DefaultTypeIndex][type.Value - 1] = type.Key;
            }

            dt.Rows[DefaultHeaderIndex][0] = ConHeaderRfinstrumentsetup;
            dt.Rows[DefaultHeaderIndex][1] = ConHeaderSubsetting;
            dt.Rows[DefaultHeaderIndex][2] = ConHeaderMeastype;
            dt.Rows[DefaultHeaderIndex][3] = ConHeaderInstrumenttype;

            foreach (InstrumentTypePara typePara in TypeParameterHeaderList)
            {
                foreach (KeyValuePair<string, int> para in typePara.DicPara)
                {
                    dt.Rows[DefaultHeaderIndex][para.Value - 1] = para.Key;
                }
            }

            foreach (KeyValuePair<int, string> item in DefaultValue)
            {
                dt.Rows[DefaultValueIndex][item.Key - 1] = item.Value;
            }
            return dt;
        }

        public void WriteRowData(InstrumentSetupRow instrumentSetupRow, ref int rowNumber, List<string> tmpInstrumentTypeList, Dictionary<string, string> dicInfoPara, string insType, int subsetIndex = 0)
        {

            string subSettingName = "";
            InitializeRow(instrumentSetupRow, rowNumber, insType);

            foreach (string instType in tmpInstrumentTypeList)
            {
                subSettingName = ProcessInstrumentTypeRow(instrumentSetupRow, rowNumber, instType, dicInfoPara, subsetIndex, subSettingName);
            }

            subSettingName = ApplyDracoSubSetting(instrumentSetupRow, rowNumber, subSettingName);
            instrumentSetupRow.SubSetting = subSettingName;
            instrumentSetupRow.InstrumentType = tmpInstrumentTypeList;
        }

        private void InitializeRow(InstrumentSetupRow instrumentSetupRow, int rowNumber, string insType)
        {
            instrumentSetupRow.InstrumentDataTable.Rows.Add(instrumentSetupRow.InstrumentDataTable.NewRow());
            instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][_headerIndices.RfInstrumentSetupIndex] = instrumentSetupRow.SetupName;
            instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][_headerIndices.InstrumentTypeIndex] = insType;
            instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][_headerIndices.MeasTypeIndex] = instrumentSetupRow.MeasSeqType;
        }

        private string ProcessInstrumentTypeRow(InstrumentSetupRow instrumentSetupRow, int rowNumber, string instType, Dictionary<string, string> dicInfoPara, int subsetIndex, string subSettingName)
        {
            string type = instType;
            InstrumentTypePara? targetType =
                TypeParameterHeaderList.FirstOrDefault(
                    x => x.InstrumentType.EqualsIgnoreCase(type));
            if (type.EqualsIgnoreCase("Error Frequency") ||
                type.EqualsIgnoreCase("Error Pin Type"))
            {
                subSettingName = "XXX_" + instrumentSetupRow.SeqIndex + "_" + subsetIndex;
                instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][_headerIndices.SubSettingIndex] = subSettingName;
            }
            if (targetType != null)
            {
                subSettingName = InstrumentSetupWriterHelpers.GetSubSettingName(type) + "_" + instrumentSetupRow.SeqIndex + "_" + subsetIndex;
                instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][_headerIndices.SubSettingIndex] = subSettingName;
                InstrumentSetupWriterHelpers.WriteParaValues(instrumentSetupRow, rowNumber, dicInfoPara, targetType);

                WritePathValue(instrumentSetupRow, rowNumber, instType, targetType);

                if (!WritePinNameValue(instrumentSetupRow, rowNumber, type, instType, targetType))
                {
                    return subSettingName;
                }

                //else if (targetType.DicPara.ContainsKey("pinlist"))
                //{
                //    if (instType == "LitePoint Source")
                //        rowData.InstrumentDataTable.Rows[rowNumber][targetType.DicPara["pinname"] - 1] = "IFSRC_MEAS1";
                //    else
                //        rowData.InstrumentDataTable.Rows[rowNumber][targetType.DicPara["pinlist"] - 1] = rowData.Pins;
                //}

                InstrumentSetupWriterHelpers.WriteModeValue(instrumentSetupRow, rowNumber, targetType);
                // new setup for LookBack #33,#34
                InstrumentSetupWriterHelpers.WritePathCtrValue(instrumentSetupRow, rowNumber, dicInfoPara, targetType);
            }
            return subSettingName;
        }

        private void WritePathValue(InstrumentSetupRow instrumentSetupRow, int rowNumber, string instType, InstrumentTypePara instrumentTypePara)
        {
            if (instrumentTypePara.DicPara.ContainsKey("path"))
            {
                InstrumentTypeSetting? instrumentTypeSetting = InstTypeSetting.FirstOrDefault(p => p.InstrumentType == instType);
                if (instrumentTypeSetting != null)
                {
                    //var instPath = instrumentTypeSetting.Path;
                    if (instrumentSetupRow.Pins.EndsWithIgnoreCase("_UW"))
                    {
                        //for issue #53 remove duplicate _UW
                        instrumentSetupRow.Pins = instrumentSetupRow.Pins[..^"_UW".Length];
                    }
                    //var path = rowData.Pins.Split(new[] { "::", "," }, StringSplitOptions.None)[0] + "-" + instrumentTypeSetting.Path;
                    //string path = string.Join(",", rowData.Pins.Split(',').Select(p => p.Split(new[] { "::" }, StringSplitOptions.None)[0] + "-" + instrumentTypeSetting.Path));
                    string path = string.Join(",", instrumentSetupRow.Pins.Split(',').Select(pinGroup => { string pPin = pinGroup.Split(["::"], StringSplitOptions.None).First(); pPin = pPin.EndsWith("_SRC") ? pPin[..^"_SRC".Length] : pPin; return $"{pPin}-{instrumentTypeSetting.Path}"; }));
                    //Replace "_LX" in each LX pin name
                    path = string.Join(",", path.Split(',').Select(x => x.Replace("_LX-", "-")));
                    //var path = rowData.Pins.Split(',').Select(pin =>
                    //    pin.Split(new[] { "::" }, StringSplitOptions.None)[0] + "-" + instrumentTypeSetting.Path).ToList();
                    //var path = rowData.Pins.Split(',').Select(pin => pin + "-" + instrumentTypeSetting.Path).ToList();

                    ReplaceInstrumentClass? needreplacekey = ReplaceInstKey.Find(cls => cls.Parameter.EqualsIgnoreCase("path") &&
                                            cls.InstrumentType.EqualsIgnoreCase(instrumentTypeSetting.InstrumentType));
                    if (needreplacekey != null)
                    {
                        foreach (KeyValuePair<string, string> pair in needreplacekey.ReplaceKey)
                        {
                            path = path.Replace(pair.Key, pair.Value);
                        }
                    }

                    instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][instrumentTypePara.DicPara["path"] - 1] = path;
                    //rowData.InstrumentDataTable.Rows[rowNumber][targetType.DicPara["path"] - 1] = string.Join(",", path);
                    //rowData.InstrumentDataTable.Rows[rowNumber][targetType.DicPara["path"] - 1] = (rowData.Pin.PinName + "-" + instPath).TrimEnd('-');
                }
            }
        }

        private bool WritePinNameValue(InstrumentSetupRow instrumentSetupRow, int rowNumber, string type, string instType, InstrumentTypePara instrumentTypePara)
        {
            if (instrumentTypePara.DicPara.ContainsKey("pinname") || instrumentTypePara.DicPara.ContainsKey("pinlist"))
            {
                string key = instrumentTypePara.DicPara.ContainsKey("pinname") ? "pinname" : "pinlist";
                InstrumentTypeSetting? instrumentTypeSetting = InstTypeSetting.FirstOrDefault(p => p.InstrumentType == instType);
                if (instrumentSetupRow.Pins.Length != 0)
                {
                    if (instrumentTypeSetting == null)
                    {
                        return false;
                    }

                    string pinsuffix = InstrumentSetupWriterHelpers.GetInstrSuffix(type);
                    var eachpinName = instrumentSetupRow.Pins.Split(',')
                        .Select(p => InstrumentSetupWriterHelpers.ProcessPinName(p, instrumentTypeSetting.PinExtraName, pinsuffix))
                        .ToList();
                    string pinname = string.Join(",", eachpinName);

                    ReplaceInstrumentClass? needreplacekey = ReplaceInstKey.Find(cls => cls.Parameter.EqualsIgnoreCase(key) &&
                                            cls.InstrumentType.EqualsIgnoreCase(instrumentTypeSetting.InstrumentType));
                    if (needreplacekey != null)
                    {
                        foreach (KeyValuePair<string, string> pair in needreplacekey.ReplaceKey)
                        {
                            pinname = pinname.Replace(pair.Key, pair.Value);
                        }
                    }

                    instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][instrumentTypePara.DicPara[key] - 1] = pinname;
                }
                else
                {
                    if (instrumentTypeSetting == null)
                    {
                        return false;
                    }

                    instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][instrumentTypePara.DicPara[key] - 1] = instrumentTypeSetting.PinExtraName;
                }
            }
            return true;
        }

        private string ApplyDracoSubSetting(InstrumentSetupRow instrumentSetupRow, int rowNumber, string subSettingName)
        {
            // 260318 add subsetting TBD_X_0 for draco
            if (instrumentSetupRow.MeasSeqType == "N")
            {
                subSettingName = "TBD_" + instrumentSetupRow.SeqIndex + "_0";
                instrumentSetupRow.InstrumentDataTable.Rows[rowNumber][_headerIndices.SubSettingIndex] = subSettingName;
            }
            return subSettingName;
        }

        public List<string> GetInstrumentType(MeasPin measPin, List<InstrumentTypeSetting> instrumentTypeSettings)
        {
            return _typeSelector.GetInstrumentType(measPin, instrumentTypeSettings);
        }

        public List<InstrumentTypeSetting> CheckTypeByFreqNew(string measType, Dictionary<string, string> dicPara, string pinname)
        {
            return InstrumentTypeSelector.CheckTypeByFreqNew(measType, dicPara, InstTypeSetting);
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                return true;
            }
            return false;
        }

        private List<InstrumentTypeSetting> ReadInstrumentTypeSetting()
        {
            var instTypeSetting = new List<InstrumentTypeSetting>();

            int startInstrumentRow = 0;

            while (startInstrumentRow <= EndInstrumentTypeSettingNumber)
            {
                if (InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, startInstrumentRow, 1).Trim()
.EqualsIgnoreCase("Instrument Type"))
                {
                    startInstrumentRow++;
                    break;
                }
                startInstrumentRow++;
            }

            #region Marked get row index
            //int rowIndex = 1;
            //int InstrumentTypeIndex = 1;
            //int FrequencyHighIndex = 2;
            //int FrequencyLowIndex = 3;
            //int PinCheckIndex = 4;
            //int PinExtraNameIndex = 5;
            //int PathIndex = 6;
            //int PinTypeIndex = 7;
            //while (rowIndex <= 20)
            //{
            //    if (InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, StartInstrumentRow, 1).Trim()
            //        .Equals("Instrument Type", StringComparison.CurrentCultureIgnoreCase))
            //        InstrumentTypeIndex = rowIndex;
            //    else if (InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, StartInstrumentRow, 1).Trim()
            //        .Equals("Frequency(High)", StringComparison.CurrentCultureIgnoreCase))
            //        FrequencyHighIndex = rowIndex;
            //    else if (InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, StartInstrumentRow, 1).Trim()
            //        .Equals("Frequency(Low)", StringComparison.CurrentCultureIgnoreCase))
            //        FrequencyLowIndex = rowIndex;
            //    else if (InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, StartInstrumentRow, 1).Trim()
            //        .Equals("Pin Check", StringComparison.CurrentCultureIgnoreCase))
            //        PinCheckIndex = rowIndex;
            //    else if (InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, StartInstrumentRow, 1).Trim()
            //        .Equals("Pin Extra Name", StringComparison.CurrentCultureIgnoreCase))
            //        PinExtraNameIndex = rowIndex;
            //    else if (InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, StartInstrumentRow, 1).Trim()
            //        .Equals("Path", StringComparison.CurrentCultureIgnoreCase))
            //        PathIndex = rowIndex;
            //    else if (InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, StartInstrumentRow, 1).Trim()
            //        .Equals("Pin Type", StringComparison.CurrentCultureIgnoreCase))
            //        PinTypeIndex = rowIndex;
            //}
            #endregion

            for (int j = startInstrumentRow; ; j++)
            {
                string type = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, j, 1).Trim();
                if (type.Length == 0)
                {
                    break;
                }

                string high = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, j, 2).Trim();
                string low = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, j, 3).Trim();
                string pinCheck = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, j, 4).Trim();
                string pinExtraName = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, j, 5).Trim();
                string path = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, j, 6).Trim();
                string pinType = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, j, 7).Trim();

                instTypeSetting.Add(new InstrumentTypeSetting
                {
                    InstrumentType = type,
                    HighLimit = InstrumentTypeUtility.ConvertToFreqValue(high),
                    LowLimit = InstrumentTypeUtility.ConvertToFreqValue(low),
                    PinExtraName = pinExtraName,
                    PinCheck = pinCheck,
                    Path = path,
                    PinType = pinType
                });
            }

            //ReplaceInstrumSheet

            return instTypeSetting;
        }

        private SortedList<int, string> ReadDefaultValue()
        {
            var defaultValue = new SortedList<int, string>();
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                defaultValue.Add(i, InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, 3, i).Trim());
            }

            //row = new InstrumentRow();
            //row.RowNum = i;
            //if (_rfinstrumentsetupIndex != -1)
            //    row.Rfinstrumentsetup = ExcelOperation.GetMergerdCellValue(_excelWorksheet, i, _rfinstrumentsetupIndex).Trim();
            //if (_subsettingIndex != -1)
            //    row.Subsetting = ExcelOperation.GetMergerdCellValue(_excelWorksheet, i, _subsettingIndex).Trim();
            //if (_meastypeIndex != -1)
            //    row.Meastype = ExcelOperation.GetMergerdCellValue(_excelWorksheet, i, _meastypeIndex).Trim();
            //if (_instrumenttypeIndex != -1)
            //    row.Instrumenttype = ExcelOperation.GetMergerdCellValue(_excelWorksheet, i, _instrumenttypeIndex).Trim();

            //foreach (var type in _typeIndexList)
            //{
            //    var targetType = row.GetTypeDictionary(type.Key);
            //    var targetHeader = _typeParameterHeaderList.FirstOrDefault(x => x.InstrumentType.Equals(type.Key));
            //    var index = type.Value;

            //    if (targetHeader == null) return null;
            //    foreach (var para in targetHeader.DicPara)
            //    {
            //        var value = ExcelOperation.GetCellValue(_excelWorksheet, i, index);
            //        targetType.Add(para.Key, value);
            //        index++;
            //    }
            //}
            //}
            return defaultValue;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = InstrumentTypeUtility.GetCellValue(_excelWorksheet, 2, i).Trim();
                if (lStrHeader.EqualsIgnoreCase(ConHeaderRfinstrumentsetup))
                {
                    _headerIndices.RfInstrumentSetupIndex = i - 1;
                    _instrumentSheet.HeaderIndex.Add(ConHeaderRfinstrumentsetup, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderSubsetting))
                {
                    _headerIndices.SubSettingIndex = i - 1;
                    _instrumentSheet.HeaderIndex.Add(ConHeaderSubsetting, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderMeastype))
                {
                    _headerIndices.MeasTypeIndex = i - 1;
                    _instrumentSheet.HeaderIndex.Add(ConHeaderMeastype, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderInstrumenttype))
                {
                    _headerIndices.InstrumentTypeIndex = i - 1;
                    _instrumentSheet.HeaderIndex.Add(ConHeaderInstrumenttype, i);
                    continue;
                }

                if (!string.IsNullOrEmpty(lStrHeader))
                { }
            }

            foreach (KeyValuePair<string, int> header in _instrumentSheet.HeaderIndex)
            {
                if (header.Value == -1)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
