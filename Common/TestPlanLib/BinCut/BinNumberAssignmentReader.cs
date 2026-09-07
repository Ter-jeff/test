using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public class BinNumberAssignmentReader : MySheetReader<BinNumberAssignmentSheet>
    {
        private const string ConHeaderBinned = "Binned";
        private const string ConHeaderDomain = "Domain";
        private const string ConHeaderMode = "Mode";
        private const string ConHeaderBinningfail1 = "Binning Fail";
        private const string ConHeaderLvccfail1 = "LVCC Fail";
        private const string ConHeaderBinningfail2 = "Binning Fail";
        private const string ConHeaderLvccfail2 = "LVCC Fail";
        private const string ConHeaderBinningfail3 = "Binning Fail";
        private const string ConHeaderLvccfail3 = "LVCC Fail";
        private const string ConHeaderBinningfail4 = "Binning Fail";
        private const string ConHeaderLvccfail4 = "LVCC Fail";
        private const string ConHeaderBinningfail5 = "Binning Fail";
        private const string ConHeaderLvccfail5 = "LVCC Fail";
        private const string ConHeaderBinningfail6 = "Binning Fail";
        private const string ConHeaderLvccfail6 = "LVCC Fail";

        private int _indexBinned = -1;
        private int _indexDomain = -1;
        private int _indexMode = -1;
        private int _indexBinningfail1 = -1;
        private int _indexLvccfail1 = -1;
        private int _indexBinningfail2 = -1;
        private int _indexLvccfail2 = -1;
        private int _indexBinningfail3 = -1;
        private int _indexLvccfail3 = -1;
        private int _indexBinningfail4 = -1;
        private int _indexLvccfail4 = -1;
        private int _indexBinningfail5 = -1;
        private int _indexLvccfail5 = -1;
        private int _indexBinningfail6 = -1;
        private int _indexLvccfail6 = -1;

        public override BinNumberAssignmentSheet? ReadSheet(ExcelWorksheet excelWorksheet)
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

            BinNumberAssignmentSheet binNumberAssignmentSheet = ReadSheetData(sheetName);

            return binNumberAssignmentSheet;
        }

        private BinNumberAssignmentSheet ReadSheetData(string sheetName)
        {
            var binNumberAssignmentSheet = new BinNumberAssignmentSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new BinNumberAssignmentRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexBinned != -1)
                {
                    row.Binned = ExcelWorksheet.GetCellValue(i, _indexBinned).Trim();
                }

                if (_indexDomain != -1)
                {
                    row.Domain = ExcelWorksheet.GetCellValue(i, _indexDomain).Trim();
                }

                if (_indexMode != -1)
                {
                    row.Mode = ExcelWorksheet.GetCellValue(i, _indexMode).Trim();
                }

                if (_indexBinningfail1 != -1)
                {
                    row.Binningfail1 = ExcelWorksheet.GetCellValue(i, _indexBinningfail1).Trim();
                }

                if (_indexLvccfail1 != -1)
                {
                    row.Lvccfail1 = ExcelWorksheet.GetCellValue(i, _indexLvccfail1).Trim();
                }

                if (_indexBinningfail2 != -1)
                {
                    row.Binningfail2 = ExcelWorksheet.GetCellValue(i, _indexBinningfail2).Trim();
                }

                if (_indexLvccfail2 != -1)
                {
                    row.Lvccfail2 = ExcelWorksheet.GetCellValue(i, _indexLvccfail2).Trim();
                }

                if (_indexBinningfail3 != -1)
                {
                    row.Binningfail3 = ExcelWorksheet.GetCellValue(i, _indexBinningfail3).Trim();
                }

                if (_indexLvccfail3 != -1)
                {
                    row.Lvccfail3 = ExcelWorksheet.GetCellValue(i, _indexLvccfail3).Trim();
                }

                if (_indexBinningfail4 != -1)
                {
                    row.Binningfail4 = ExcelWorksheet.GetCellValue(i, _indexBinningfail4).Trim();
                }

                if (_indexLvccfail4 != -1)
                {
                    row.Lvccfail4 = ExcelWorksheet.GetCellValue(i, _indexLvccfail4).Trim();
                }

                if (_indexBinningfail5 != -1)
                {
                    row.Binningfail5 = ExcelWorksheet.GetCellValue(i, _indexBinningfail5).Trim();
                }

                if (_indexLvccfail5 != -1)
                {
                    row.Lvccfail5 = ExcelWorksheet.GetCellValue(i, _indexLvccfail5).Trim();
                }

                if (_indexBinningfail6 != -1)
                {
                    row.Binningfail6 = ExcelWorksheet.GetCellValue(i, _indexBinningfail6).Trim();
                }

                if (_indexLvccfail6 != -1)
                {
                    row.Lvccfail6 = ExcelWorksheet.GetCellValue(i, _indexLvccfail6).Trim();
                }

                binNumberAssignmentSheet.Rows.Add(row);
            }

            binNumberAssignmentSheet.IndexBinned = _indexBinned;
            binNumberAssignmentSheet.IndexDomain = _indexDomain;
            binNumberAssignmentSheet.IndexMode = _indexMode;
            binNumberAssignmentSheet.IndexBinningfail1 = _indexBinningfail1;
            binNumberAssignmentSheet.IndexLvccfail1 = _indexLvccfail1;
            binNumberAssignmentSheet.IndexBinningfail2 = _indexBinningfail2;
            binNumberAssignmentSheet.IndexLvccfail2 = _indexLvccfail2;
            binNumberAssignmentSheet.IndexBinningfail3 = _indexBinningfail3;
            binNumberAssignmentSheet.IndexLvccfail3 = _indexLvccfail3;
            binNumberAssignmentSheet.IndexBinningfail4 = _indexBinningfail4;
            binNumberAssignmentSheet.IndexLvccfail4 = _indexLvccfail4;
            binNumberAssignmentSheet.IndexBinningfail5 = _indexBinningfail5;
            binNumberAssignmentSheet.IndexLvccfail5 = _indexLvccfail5;
            binNumberAssignmentSheet.IndexBinningfail6 = _indexBinningfail6;
            binNumberAssignmentSheet.IndexLvccfail6 = _indexLvccfail6;

            return binNumberAssignmentSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string h3 = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                string h2 = StartRow - 1 > 0 ? ExcelWorksheet.GetCellValue(StartRow - 1, i).Trim() : "";
                string h1 = StartRow - 2 > 0 ? ExcelWorksheet.GetCellValue(StartRow - 2, i).Trim() : "";

                if (h3.EqualsIgnoreCase(ConHeaderBinned))
                {
                    _indexBinned = i;
                }
                else if (h3.EqualsIgnoreCase(ConHeaderDomain))
                {
                    _indexDomain = i;
                }
                else if (h3.EqualsIgnoreCase(ConHeaderMode))
                {
                    _indexMode = i;
                }

                // Composite Matches (Category + Type)
                else if (IsMatch(h1, h2, h3, "TD", "Softbin", ConHeaderBinningfail1))
                {
                    _indexBinningfail1 = i;
                }
                else if (IsMatch(h1, h2, h3, "TD", "Softbin", ConHeaderLvccfail1))
                {
                    _indexLvccfail1 = i;
                }
                else if (IsMatch(h1, h2, h3, "TD", "HardBin", ConHeaderBinningfail2))
                {
                    _indexBinningfail2 = i;
                }
                else if (IsMatch(h1, h2, h3, "TD", "HardBin", ConHeaderLvccfail2))
                {
                    _indexLvccfail2 = i;
                }
                else if (IsMatch(h1, h2, h3, "Mbist", "Softbin", ConHeaderBinningfail3))
                {
                    _indexBinningfail3 = i;
                }
                else if (IsMatch(h1, h2, h3, "Mbist", "Softbin", ConHeaderLvccfail3))
                {
                    _indexLvccfail3 = i;
                }
                else if (IsMatch(h1, h2, h3, "Mbist", "HardBin", ConHeaderBinningfail4))
                {
                    _indexBinningfail4 = i;
                }
                else if (IsMatch(h1, h2, h3, "Mbist", "HardBin", ConHeaderLvccfail4))
                {
                    _indexLvccfail4 = i;
                }
                else if (IsMatch(h1, h2, h3, "SPI", "Softbin", ConHeaderBinningfail5))
                {
                    _indexBinningfail5 = i;
                }
                else if (IsMatch(h1, h2, h3, "SPI", "Softbin", ConHeaderLvccfail5))
                {
                    _indexLvccfail5 = i;
                }
                else if (IsMatch(h1, h2, h3, "SPI", "HardBin", ConHeaderBinningfail6))
                {
                    _indexBinningfail6 = i;
                }
                else if (IsMatch(h1, h2, h3, "SPI", "HardBin", ConHeaderLvccfail6))
                {
                    _indexLvccfail6 = i;
                }
            }
        }

        private static bool IsMatch(string h1, string h2, string h3, string expected1, string expected2, string expected3)
        {
            return h1.EqualsIgnoreCase(expected1) &&
                   h2.EqualsIgnoreCase(expected2) &&
                   h3.EqualsIgnoreCase(expected3);
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderBinned))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class BinNumberAssignmentSheet : MySheet
    {
        public List<BinNumberAssignmentRow> Rows { get; }

        public int IndexBinned = -1;
        public int IndexDomain = -1;
        public int IndexMode = -1;
        public int IndexBinningfail1 = -1;
        public int IndexLvccfail1 = -1;
        public int IndexBinningfail2 = -1;
        public int IndexLvccfail2 = -1;
        public int IndexBinningfail3 = -1;
        public int IndexLvccfail3 = -1;
        public int IndexBinningfail4 = -1;
        public int IndexLvccfail4 = -1;
        public int IndexBinningfail5 = -1;
        public int IndexLvccfail5 = -1;
        public int IndexBinningfail6 = -1;
        public int IndexLvccfail6 = -1;

        #region Constructor
        public BinNumberAssignmentSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion
    }

    public class BinNumberAssignmentRow : MyRow
    {
        public string Binned { get; set; }
        public string Domain { get; set; }
        public string Mode { get; set; }
        public string Binningfail1 { get; set; }
        public string Lvccfail1 { get; set; }
        public string Binningfail2 { get; set; }
        public string Lvccfail2 { get; set; }
        public string Binningfail3 { get; set; }
        public string Lvccfail3 { get; set; }
        public string Binningfail4 { get; set; }
        public string Lvccfail4 { get; set; }
        public string Binningfail5 { get; set; }
        public string Lvccfail5 { get; set; }
        public string Binningfail6 { get; set; }
        public string Lvccfail6 { get; set; }

        #region Constructor
        public BinNumberAssignmentRow(string sheetName = "")
        {
            SheetName = sheetName;
            Binned = "";
            Domain = "";
            Mode = "";
            Binningfail1 = "";
            Lvccfail1 = "";
            Binningfail2 = "";
            Lvccfail2 = "";
            Binningfail3 = "";
            Lvccfail3 = "";
            Binningfail4 = "";
            Lvccfail4 = "";
            Binningfail5 = "";
            Lvccfail5 = "";
            Binningfail6 = "";
            Lvccfail6 = "";
        }
        #endregion
    }
}
