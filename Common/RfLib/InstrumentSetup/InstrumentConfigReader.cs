using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using CommonLib.Extension;

using OfficeOpenXml;

using RfLib.InstrumentSetup.InstrumentTypeData;

namespace RfLib.InstrumentSetup
{
    [ExcludeFromCodeCoverage]
    public class InstrumentConfigReader
    {
        private const string ConHeaderRfinstrumentsetup = "RF Instrument Setup";
        private const string ConHeaderSubsetting = "SubSetting";
        private const string ConHeaderMeastype = "MeasType";
        private const string ConHeaderInstrumenttype = "Instrument Type";

        private ExcelWorksheet _excelWorksheet = null!;
        private string _sheetName = "";
        private InstrumentConfigSheet _instrumentSheet = new("");

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _rfinstrumentsetupIndex = -1;
        private int _subsettingIndex = -1;
        private int _meastypeIndex = -1;
        private int _instrumenttypeIndex = -1;

        private readonly Dictionary<string, bool> _headerOptional = new()
        {
            {
                "RF Instrument Setup",true
            },
            {
                "SubSetting",true
            },
            {
                "MeasType",true
            },
            {
                "Instrument Type",true
            }
        };

        private Dictionary<string, int> _typeIndexList = [];
        private List<InstrumentTypePara> _typeParameterHeaderList = [];

        public InstrumentConfigSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null!;
            }

            _excelWorksheet = excelWorksheet;

            _sheetName = excelWorksheet.Name;

            _instrumentSheet = new InstrumentConfigSheet(_sheetName);

            Reset();

            if (!GetDimensions())
            {
                return null!;
            }

            if (!GetHeaderIndex())
            {
                return null!;
            }

            _typeIndexList = InstrumentTypeUtility.GetInstrumentType(excelWorksheet, _endColNumber);
            _typeParameterHeaderList = InstrumentTypeUtility.GetInstrumentParameterHeader(excelWorksheet, _typeIndexList, _endColNumber);
            _instrumentSheet = ReadSheetData();

            return _instrumentSheet;
        }

        private InstrumentConfigSheet ReadSheetData()
        {
            InstrumentConfigRow row;
            for (int i = _startRowNumber + 2; i <= _endRowNumber; i++)
            {
                row = new InstrumentConfigRow { RowNum = i };
                if (_rfinstrumentsetupIndex != -1)
                {
                    row.Rfinstrumentsetup = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, i, _rfinstrumentsetupIndex).Trim();
                }

                if (_subsettingIndex != -1)
                {
                    row.Subsetting = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, i, _subsettingIndex).Trim();
                }

                if (_meastypeIndex != -1)
                {
                    row.Meastype = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, i, _meastypeIndex).Trim();
                }

                if (_instrumenttypeIndex != -1)
                {
                    row.Instrumenttype = InstrumentTypeUtility.GetMergerdCellValue(_excelWorksheet, i, _instrumenttypeIndex).Trim();
                }

                foreach (KeyValuePair<string, int> type in _typeIndexList)
                {
                    Dictionary<string, string> targetType = row.GetTypeDictionary(type.Key);
                    InstrumentTypePara? targetHeader =
                        _typeParameterHeaderList.FirstOrDefault(x => x.InstrumentType == type.Key);
                    int index = type.Value;

                    if (targetHeader == null)
                    {
                        return null!;
                    }

                    foreach (KeyValuePair<string, int> para in targetHeader.DicPara)
                    {
                        string value = InstrumentTypeUtility.GetCellValue(_excelWorksheet, i, index);
                        targetType.Add(para.Key, value);
                        index++;
                    }
                }

                _instrumentSheet.Rows.Add(row);
            }
            return _instrumentSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = InstrumentTypeUtility.GetCellValue(_excelWorksheet, 2, i).Trim();
                if (lStrHeader.EqualsIgnoreCase(ConHeaderRfinstrumentsetup))
                {
                    _rfinstrumentsetupIndex = i;
                    _instrumentSheet.HeaderIndex.Add(ConHeaderRfinstrumentsetup, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderSubsetting))
                {
                    _subsettingIndex = i;
                    _instrumentSheet.HeaderIndex.Add(ConHeaderSubsetting, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderMeastype))
                {
                    _meastypeIndex = i;
                    _instrumentSheet.HeaderIndex.Add(ConHeaderMeastype, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderInstrumenttype))
                {
                    _instrumenttypeIndex = i;
                    _instrumentSheet.HeaderIndex.Add(ConHeaderInstrumenttype, i);
                    continue;
                }

                if (!string.IsNullOrEmpty(lStrHeader))
                { }
            }

            foreach (KeyValuePair<string, int> header in _instrumentSheet.HeaderIndex)
            {
                if (header.Value == -1 && _headerOptional.TryGetValue(header.Key, out bool value) && value)
                {
                    return false;
                }
            }

            return true;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }
            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _rfinstrumentsetupIndex = -1;
            _subsettingIndex = -1;
            _meastypeIndex = -1;
            _instrumenttypeIndex = -1;
        }
    }

    [ExcludeFromCodeCoverage]
    public class InstrumentConfigSheet(string sheetname)
    {
        #region Properity
        public string SheetName { get; set; } = sheetname;
        public List<InstrumentConfigRow> Rows { get; } = [];
        public Dictionary<string, int> HeaderIndex { get; } = [];

        #endregion
        #region Constructor
        #endregion

        public InstrumentConfigRow GetInstrumentType(string setupKey)
        {
            if (string.IsNullOrEmpty(setupKey))
            {
                return null!;
            }

            string majorSetup = setupKey.Split('#')[0];
            string minorSetup = setupKey.Split('#')[1];
            return Rows.FirstOrDefault(p => p.Rfinstrumentsetup == majorSetup && p.Subsetting == minorSetup)!;
        }
    }

    [ExcludeFromCodeCoverage]
    public class InstrumentConfigRow
    {
        #region Properity
        public string SourceSheetName { set; get; }

        public int RowNum { get; set; }

        public string Rfinstrumentsetup { set; get; }

        public string Subsetting { set; get; }

        public string Meastype { set; get; }

        public string Instrumenttype { set; get; }

        public Dictionary<string, string> LxPlusSource;
        public Dictionary<string, string> LxPlusAnalyzer;
        public Dictionary<string, string> LitePointSource;
        public Dictionary<string, string> LitePointAnalyzer;
        public Dictionary<string, string> UltraWave24CwSource;
        public Dictionary<string, string> UltraWave24MultiToneSource;
        public Dictionary<string, string> UltraWave24ModulatedSource;
        public Dictionary<string, string> UltraWave24Receiver;
        public Dictionary<string, string> UltraPac80Source;
        public Dictionary<string, string> UltraPac80Capture;
        public Dictionary<string, string> UltraPac80Vdm;
        public Dictionary<string, string> Gigadig;
        public Dictionary<string, string> Uvi80Vdm;
        public Dictionary<string, string> Uvi80Meter;
        public Dictionary<string, string> Uvs256Meter;
        public Dictionary<string, string> Uvi80Source;
        public Dictionary<string, string> Up1600;
        public Dictionary<string, string> Mw7GSource;
        public Dictionary<string, string> Mw7GAnalyzer;
        public Dictionary<string, string> SwitchControl;
        #endregion

        #region Constructor

        public InstrumentConfigRow()
        {
            SourceSheetName = "";
            Rfinstrumentsetup = "";
            Subsetting = "";
            Meastype = "";
            Instrumenttype = "";

            LxPlusSource = [];
            LxPlusAnalyzer = [];
            LitePointSource = [];
            LitePointAnalyzer = [];
            UltraWave24CwSource = [];
            UltraWave24MultiToneSource = [];
            UltraWave24ModulatedSource = [];
            UltraWave24Receiver = [];
            UltraPac80Source = [];
            UltraPac80Capture = [];
            UltraPac80Vdm = [];
            Gigadig = [];
            Uvi80Vdm = [];
            Uvi80Meter = [];
            Uvs256Meter = [];
            Uvi80Source = [];
            Up1600 = [];
            Mw7GSource = [];
            Mw7GAnalyzer = [];
            SwitchControl = [];
        }

        public InstrumentConfigRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Rfinstrumentsetup = "";
            Subsetting = "";
            Meastype = "";
            Instrumenttype = "";

            LxPlusSource = [];
            LxPlusAnalyzer = [];
            LitePointSource = [];
            LitePointAnalyzer = [];
            UltraWave24CwSource = [];
            UltraWave24MultiToneSource = [];
            UltraWave24ModulatedSource = [];
            UltraWave24Receiver = [];
            UltraPac80Source = [];
            UltraPac80Capture = [];
            UltraPac80Vdm = [];
            Gigadig = [];
            Uvi80Vdm = [];
            Uvi80Meter = [];
            Uvs256Meter = [];
            Uvi80Source = [];
            Up1600 = [];
            Mw7GSource = [];
            Mw7GAnalyzer = [];
            SwitchControl = [];
        }

        public Dictionary<string, string> GetTypeDictionary(string type)
        {
            return type switch
            {
                "LXPlus Source" => LxPlusSource,
                "LXPlus Analyzer" => LxPlusAnalyzer,
                "LitePoint Source" => LitePointSource,
                "LitePoint Analyzer" => LitePointAnalyzer,
                "UltraWave24 CW Source" => UltraWave24CwSource,
                "UltraWave24 MultiTone Source" => UltraWave24MultiToneSource,
                "UltraWave24 Modulated Source" => UltraWave24ModulatedSource,
                "UltraWave24 Receiver" => UltraWave24Receiver,
                "UltraPac80 Source" => UltraPac80Source,
                "UltraPac80 Capture" => UltraPac80Capture,
                "UltraPac80 Vdm" => UltraPac80Vdm,
                "Gigadig" => Gigadig,
                "UVI80 Vdm" => Uvi80Vdm,
                "UVI80 Meter" => Uvi80Meter,
                "UVS256 Meter" => Uvs256Meter,
                "UVI80 Source" => Uvi80Source,
                "UP1600" => Up1600,
                "MW7G Source" => Mw7GSource,
                "MW7G Analyzer" => Mw7GAnalyzer,
                "SWITCH CONTROL" => SwitchControl,
                _ => null!,
            };
        }
        #endregion
    }
}
