using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Harvest
{
    public class DigitalFlagsReader : MySheetReader<DigitalFlagsSheet>
    {
        private const string ConHeaderIp = "IP";
        private const string ConHeaderFlagName = "Flag Name";
        private const string ConHeaderStatement = "Statement";
        private const string ConHeaderPrint = "Print";
        private const string ConHeaderScanFlag = "SCAN Flag";
        private const string ConHeaderMbistFlag = "MBIST Flag";
        private const string ConHeaderComments = "Comments";

        private int _indexIp = -1;
        private int _indexFlagName = -1;
        private int _indexStatement = -1;
        private int _indexPrint = -1;
        private int _indexScanFlag = -1;
        private int _indexMbistFlag = -1;
        private int _indexComments = -1;

        public override DigitalFlagsSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            DigitalFlagsSheet digitalFlagsSheet = ReadSheet(sheetName);

            return digitalFlagsSheet;
        }

        private DigitalFlagsSheet ReadSheet(string sheetName)
        {
            var digitalFlagsSheet = new DigitalFlagsSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new DigitalFlagsSheetRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexIp != -1)
                {
                    row.Ip = ExcelWorksheet.GetCellValue(i, _indexIp).Trim();
                }

                if (_indexFlagName != -1)
                {
                    row.FlagName = ExcelWorksheet.GetCellValue(i, _indexFlagName).Trim();
                }

                if (row.FlagName?.Length == 0)
                {
                    continue;
                }

                if (_indexStatement != -1)
                {
                    row.Statement = ExcelWorksheet.GetCellValue(i, _indexStatement).Trim();
                }

                if (_indexPrint != -1)
                {
                    row.Print = ExcelWorksheet.GetCellValue(i, _indexPrint).Trim();
                }

                if (_indexScanFlag != -1)
                {
                    row.ScanFlag = ExcelWorksheet.GetCellValue(i, _indexScanFlag).Trim();
                }

                if (_indexMbistFlag != -1)
                {
                    row.MbistFlag = ExcelWorksheet.GetCellValue(i, _indexMbistFlag).Trim();
                }

                if (_indexComments != -1)
                {
                    row.Comments = ExcelWorksheet.GetCellValue(i, _indexComments).Trim();
                }

                digitalFlagsSheet.Rows.Add(row);
            }

            digitalFlagsSheet.IndexIp = _indexIp;
            digitalFlagsSheet.IndexFlagName = _indexFlagName;
            digitalFlagsSheet.IndexStatement = _indexStatement;
            digitalFlagsSheet.IndexPrint = _indexPrint;
            digitalFlagsSheet.IndexScanFlag = _indexScanFlag;
            digitalFlagsSheet.IndexMbistFlag = _indexMbistFlag;
            digitalFlagsSheet.IndexComments = _indexComments;

            return digitalFlagsSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderIp))
                {
                    _indexIp = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderFlagName))
                {
                    _indexFlagName = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderStatement))
                {
                    _indexStatement = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderPrint))
                {
                    _indexPrint = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderScanFlag))
                {
                    _indexScanFlag = i;
                }
                if (header.EqualsIgnoreCase(ConHeaderMbistFlag))
                {
                    _indexMbistFlag = i;
                }
                if (header.EqualsIgnoreCase(ConHeaderComments))
                {
                    _indexComments = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderIp))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
