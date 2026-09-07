using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Harvest
{
    public class DomainFlagsReader : MySheetReader<DomainFlagsSheet>
    {
        private const string ConHeaderIp = "IP";
        private const string ConHeaderFuseFlag = "Fuse Flag (SCAN Flag or MBIST Flag)";
        private const string ConHeaderScanFlag = "SCAN Flag";
        private const string ConHeaderMbistFlag = "MBIST Flag";

        private int _indexIp = -1;
        private int _indexFuseFlag = -1;
        private int _indexScanFlag = -1;
        private int _indexMbistFlag = -1;

        public override DomainFlagsSheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            DomainFlagsSheet domainFlagsSheet = ReadSheet(sheetName);

            return domainFlagsSheet;
        }

        private DomainFlagsSheet ReadSheet(string sheetName)
        {
            var domainFlagsSheet = new DomainFlagsSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new DomainFlagsRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexIp != -1)
                {
                    row.Ip = ExcelWorksheet.GetCellValue(i, _indexIp).Trim();
                }

                if (row.Ip?.Length == 0)
                {
                    continue;
                }

                if (_indexFuseFlag != -1)
                {
                    row.FuseFlag = ExcelWorksheet.GetCellValue(i, _indexFuseFlag).Trim();
                }

                if (_indexScanFlag != -1)
                {
                    row.ScanFlag = ExcelWorksheet.GetCellValue(i, _indexScanFlag).Trim();
                }

                if (_indexMbistFlag != -1)
                {
                    row.MbistFlag = ExcelWorksheet.GetCellValue(i, _indexMbistFlag).Trim();
                }

                domainFlagsSheet.Rows.Add(row);
            }

            domainFlagsSheet.IndexIp = _indexIp;
            domainFlagsSheet.IndexFuseFlag = _indexFuseFlag;
            domainFlagsSheet.IndexScanFlag = _indexScanFlag;
            domainFlagsSheet.IndexMbistFlag = _indexMbistFlag;

            return domainFlagsSheet;
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
                if (header.EqualsIgnoreCase(ConHeaderFuseFlag))
                {
                    _indexFuseFlag = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderScanFlag))
                {
                    _indexScanFlag = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderMbistFlag))
                {
                    _indexMbistFlag = i;
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

    public class DomainFlagsSheet : MySheet
    {
        public List<DomainFlagsRow> Rows { set; get; }

        public int IndexIp = -1;
        public int IndexFuseFlag = -1;
        public int IndexScanFlag = -1;
        public int IndexMbistFlag = -1;
        #region Constructor
        public DomainFlagsSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion

        public List<string> GetAllFlags()
        {
            var flags = new List<string>();
            foreach (DomainFlagsRow row in Rows)
            {
                if (!flags.Exists(x => x == row.FuseFlag) && !string.IsNullOrEmpty(row.FuseFlag))
                {
                    flags.Add(row.FuseFlag);
                }

                if (!flags.Exists(x => x == row.ScanFlag) && !string.IsNullOrEmpty(row.ScanFlag))
                {
                    flags.Add(row.ScanFlag);
                }

                if (!flags.Exists(x => x == row.MbistFlag) && !string.IsNullOrEmpty(row.MbistFlag))
                {
                    flags.Add(row.MbistFlag);
                }
            }
            return flags;
        }
    }

    public class DomainFlagsRow : MyRow
    {
        public string Ip = "";
        public string FuseFlag = "";
        public string ScanFlag = "";
        public string MbistFlag = "";
        #region Constructor
        public DomainFlagsRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion

    }
}
