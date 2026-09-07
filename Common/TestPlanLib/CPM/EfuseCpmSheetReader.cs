using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.CPM
{
    public class EfuseCpmSheetReader : MySheetReader<EfuseCpmSheet>
    {
        private const string ConHeaderEfuseBank = "EFUSE bank";
        private const string ConHeaderCpmEfusename2 = "CPM EFUSE name2";
        private const string ConHeaderBits = "bits";
        private const string ConHeaderFlag = "F_";
        private const string ConHeaderCpm = "CPM";
        private const string ConHeaderComment = "comment";

        private int _efuseBankCol = -1;
        private int _cpmEfuseName2Col = -1;
        private int _bitsCol = -1;
        private readonly Dictionary<string, int> _flagCol = [];
        private int _cpmCol = -1;
        private int _commentCol = -1;

        public override EfuseCpmSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition();

            GetHeaders();

            EfuseCpmSheet sheet = ReadSheet(sheetName);

            return sheet;
        }

        protected void GetFirstHeaderPosition()
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderEfuseBank))
                    {
                        StartRow = i;
                        break;
                    }
                }
            }
        }

        protected void GetHeaders()
        {
            for (int i = 1; i <= EndCol; i++)
            {
                string lStrHeader = GetCellValue(StartRow, i).ToUpper().Trim();

                if (lStrHeader.EqualsIgnoreCase(ConHeaderEfuseBank))
                {
                    _efuseBankCol = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderCpmEfusename2))
                {
                    _cpmEfuseName2Col = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderBits))
                {
                    _bitsCol = i;
                    continue;
                }
                if (lStrHeader.StartsWithIgnoreCase(ConHeaderFlag))
                {
                    _flagCol.Add(lStrHeader, i);
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderCpm))
                {
                    _cpmCol = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderComment))
                {
                    _commentCol = i;
                    continue;
                }
            }
        }

        protected EfuseCpmSheet ReadSheet(string sheetName)
        {
            var sheet = new EfuseCpmSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new EfuseCpmSheetRow
                {
                    RowNum = i,
                    SheetName = sheetName,
                    EfuseBank = _efuseBankCol == -1 ? "" : ExcelWorksheet.GetCellValue(i, _efuseBankCol).Trim(),
                    CpmEfuseName2 = _cpmEfuseName2Col == -1 ? "" : ExcelWorksheet.GetCellValue(i, _cpmEfuseName2Col).Trim(),
                    Bits = _bitsCol == -1 ? "" : ExcelWorksheet.GetCellValue(i, _bitsCol).Trim(),
                    //row.DefaultValues = _defaultValuesCol == -1 ? "" : ExcelWorksheet.GetCellValue( i, _defaultValuesCol).Trim();
                    //row.AlternateValues = _alternateValuesCol == -1 ? "" : ExcelWorksheet.GetCellValue( i, _alternateValuesCol).Trim();
                    Cpm = _cpmCol == -1 ? "" : ExcelWorksheet.GetCellValue(i, _cpmCol).Trim(),
                    Comment = _commentCol == -1 ? "" : ExcelWorksheet.GetCellValue(i, _commentCol).Trim()
                };
                foreach (KeyValuePair<string, int> flag in _flagCol)
                {
                    string flagName = flag.Key;
                    int col = flag.Value;
                    if (col == 0)
                    {
                        continue;
                    }

                    row.FlagsValue.Add(flagName, ExcelWorksheet.GetCellValue(i, col).Trim());
                }
                sheet.Rows.Add(row);
            }
            sheet.EfuseBankColNumber = _efuseBankCol;
            sheet.CpmEfuseName2ColNumber = _cpmEfuseName2Col;
            sheet.BitsColNumber = _bitsCol;
            sheet.FlagColNumber = _flagCol;
            //sheet.DefaultValuesColNumber = _defaultValuesCol;
            //sheet.AlternateValuesColNumber = _alternateValuesCol;
            sheet.CpmColNumber = _cpmCol;
            sheet.CommentColNumber = _commentCol;

            return sheet;
        }

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            object value = ExcelWorksheet.Cells[rowNumber, columnNumber].Value;
            return value?.ToString() ?? "";
        }
    }
}
