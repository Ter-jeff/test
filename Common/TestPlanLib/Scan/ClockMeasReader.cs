using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Scan
{
    public partial class ClockMeasReader
    {
        [GeneratedRegex(@"\d$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();
        private readonly string _headerType = "type";
        private readonly string _headerType2 = "init sequence";
        private readonly string _headerUnits = "units";
        private readonly string _headerMode = "mode";
        private readonly string _headerPostinitPattern = "post-init pattern";
        private readonly string _headerDigsrc = "digsrc";
        private readonly string _headerPin = "pin";
        private readonly List<string> _uslLsl = ["LSL", "USL"];

        private int _typeIndex = -1;
        private int _unitsIndex = -1;
        private int _modeIndex = -1;
        private int _patternIndex = -1;
        private int _digsrcIndex = -1;
        private readonly Dictionary<string, List<int>> _pinsDictionary = [];

        private int _startrow = 1;
        private ExcelWorksheet? _excelWorksheet;

        //type	mode	post-init pattern	units
        //pin
        public ClockMeasSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            _excelWorksheet = excelWorksheet;

            var clockMeasSheet = new ClockMeasSheet
            {
                SheetName = excelWorksheet.Name
            };

            ReadHeader();

            clockMeasSheet.Rows = ReadContent();

            return clockMeasSheet;
        }

        private void ReadHeader()
        {
            bool isSearchPin = false;
            for (int startrowIndex = 1; startrowIndex <= _excelWorksheet!.Dimension.Rows; startrowIndex++)
            {
                for (int j = 1; j <= _excelWorksheet!.Dimension.Columns; j++)
                {
                    if (string.IsNullOrEmpty(_excelWorksheet!.GetCellValue(startrowIndex, j)))
                    {
                        continue;
                    }

                    if (isSearchPin)
                    {
                        string pinName = _excelWorksheet!.GetCellValue(startrowIndex, j);
                        if (!_pinsDictionary.TryGetValue(pinName, out List<int>? value))
                        {
                            value = [];
                            _pinsDictionary.Add(pinName, value);
                        }

                        if (value.Count < 2)
                        {
                            value.Add(j);
                        }
                    }
                    if (_excelWorksheet!.GetCellValue(startrowIndex, j)
.EqualsIgnoreCase(_headerPin))
                    {
                        isSearchPin = true;
                    }
                }
                isSearchPin = false;
                for (int j = 1; j <= _excelWorksheet!.Dimension.Columns; j++)
                {
                    if (string.IsNullOrEmpty(_excelWorksheet!.GetCellValue(startrowIndex, j)))
                    {
                        continue;
                    }

                    string key = _excelWorksheet!.GetCellValue(startrowIndex, j);
                    if (key
.EqualsIgnoreCase(_headerType) ||
                        key
.EqualsIgnoreCase(_headerType2))
                    {
                        _typeIndex = j;
                    }

                    if (key
.EqualsIgnoreCase(_headerMode))
                    {
                        _modeIndex = j;
                    }

                    if (key
.EqualsIgnoreCase(_headerUnits))
                    {
                        _unitsIndex = j;
                    }

                    if (key
.EqualsIgnoreCase(_headerPostinitPattern))
                    {
                        _patternIndex = j;
                    }

                    if (key
.EqualsIgnoreCase(_headerDigsrc))
                    {
                        _digsrcIndex = j;
                    }

                    if (_uslLsl.Exists(p => p.EqualsIgnoreCase(key)))
                    {
                        break;
                    }
                }
                if (_typeIndex != -1)
                {
                    _startrow = startrowIndex;
                    break;
                }

            }
            _startrow++;
        }

        private List<ClockMeasRow> ReadContent()
        {
            List<ClockMeasRow> clockRows = [];
            for (int startrowIndex = _startrow; startrowIndex <= _excelWorksheet!.Dimension.Rows; startrowIndex++)
            {
                int startColIndex = 1;
                var clockMeasRow = new ClockMeasRow { RowNum = startrowIndex };
                clockRows.Add(clockMeasRow);
                for (startColIndex = 1; startColIndex <= _excelWorksheet!.Dimension.Columns; startColIndex++)
                {
                    string value = _excelWorksheet!.GetCellValue(startrowIndex, startColIndex);
                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    if (startColIndex == _typeIndex)
                    {
                        clockMeasRow.Type = value;
                    }

                    if (startColIndex == _modeIndex)
                    {
                        clockMeasRow.PMode = value;
                    }

                    if (startColIndex == _patternIndex)
                    {
                        clockMeasRow.Pattern = value;
                    }

                    if (startColIndex == _unitsIndex)
                    {
                        clockMeasRow.Units = value;
                    }

                    if (startColIndex == _digsrcIndex)
                    {
                        clockMeasRow.Digsrc = value;
                    }

                    if (_pinsDictionary.Values.ToList().Exists(p => p.Exists(q => q == startColIndex)))
                    {
                        foreach (KeyValuePair<string, List<int>> pinInfo in _pinsDictionary)
                        {
                            if (pinInfo.Value.Contains(startColIndex))
                            {
                                if (!clockMeasRow.PinInfo.TryGetValue(pinInfo.Key, out List<string>? value1))
                                {
                                    value1 = [];
                                    clockMeasRow.PinInfo.Add(pinInfo.Key, value1);
                                }

                                if (_regex.IsMatch(value))
                                {
                                    value1.Add(value + clockMeasRow.Units);
                                }
                                else
                                {
                                    value1.Add(value);
                                }

                                break;
                            }
                        }
                    }

                }
            }

            return clockRows;
        }
    }

    public class ClockMeasSheet : MySheet
    {
        public List<ClockMeasRow> Rows = [];
    }

    public class ClockMeasRow : MyRow
    {
        public string Pattern = "";
        public string Type = "";
        private string _pModeStr = "";
        public string ErrMsg = "";
        public string PMode
        {
            get
            {
                string result = "SA";
                List<string> typeList = [.. Type.Split(' ')];
                if (!string.IsNullOrEmpty(_pModeStr))
                {
                    result = _pModeStr;
                }
                else if (typeList.Count > 1)
                {
                    typeList.RemoveAt(0);
                    var validList = new List<string> { "BIST", "TD", "SA" };
                    foreach (string resulttype in typeList)
                    {
                        if (validList.Exists(p => p.EqualsIgnoreCase(resulttype)))
                        {
                            result = resulttype;
                        }
                    }
                }
                return result;
            }
            set { _pModeStr = value; }
        }

        public string Units = "";
        public Dictionary<string, List<string>> PinInfo = [];
        public string Digsrc = "";
        public static string GetLimit()
        {
            return "";
        }
    }
}
