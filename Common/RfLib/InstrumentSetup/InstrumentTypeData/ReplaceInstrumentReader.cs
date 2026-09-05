using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace RfLib.InstrumentSetup.InstrumentTypeData
{
    public class ReplaceInstrumentReader
    {
        private const string ConHeaderInstrumentType = "Instrument Type";
        private const string ConHeaderParameter = "Parameter";
        private const string ConHeaderOriginal = "Original";
        private const string ConHeaderChange = "Change";

        private ExcelWorksheet _excelWorksheet = null!;
        private ReplaceInstrumentSheet _replaceinstrumentSheet = new();

        private int _startColNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;

        public ReplaceInstrumentSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            _excelWorksheet = excelWorksheet;

            _replaceinstrumentSheet = new ReplaceInstrumentSheet();

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            _replaceinstrumentSheet = ReadSheetData();

            return _replaceinstrumentSheet;
        }

        private ReplaceInstrumentSheet ReadSheetData()
        {
            //var replaceinsclass = new ReplaceInstrumentClass();
            string instrumenttype = "";
            string parameter = "";
            var replacekey = new Dictionary<string, string>();

            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = InstrumentTypeUtility.GetCellValue(_excelWorksheet, 1, i).Trim();
                if (lStrHeader.EqualsIgnoreCase(ConHeaderInstrumentType))
                {
                    instrumenttype = InstrumentTypeUtility.GetCellValue(_excelWorksheet, 2, i).Trim();
                }
                else if (lStrHeader.EqualsIgnoreCase(ConHeaderParameter))
                {
                    parameter = InstrumentTypeUtility.GetCellValue(_excelWorksheet, 2, i).Trim();
                }
                else if (lStrHeader.EqualsIgnoreCase(ConHeaderOriginal))
                {
                    replacekey = [];
                    for (int j = 2; j < _endRowNumber + 1; j++)
                    {
                        string originalname = InstrumentTypeUtility.GetCellValue(_excelWorksheet, j, i).Trim();
                        string changename = InstrumentTypeUtility.GetCellValue(_excelWorksheet, j, i + 1).Trim();

                        if (originalname.Length != 0)
                        {
                            if (!replacekey.TryAdd(originalname, changename))
                            {
                                replacekey[originalname] = changename;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                else if (lStrHeader.EqualsIgnoreCase(ConHeaderChange))
                {
                    _replaceinstrumentSheet.ReplaceClass.Add(new ReplaceInstrumentClass(instrumenttype, parameter, replacekey));
                }
            }

            return _replaceinstrumentSheet;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                //_startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }
            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            //_startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
        }

    }
}
