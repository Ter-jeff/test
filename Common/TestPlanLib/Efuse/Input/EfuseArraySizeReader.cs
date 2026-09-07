using System;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Efuse.Input
{
    public partial class EfuseArraySizeReader : MySheetReader<EfuseArraySizeSheet>
    {
        private const string ConHeaderFuseArray = "^Fuse Array";
        private const string ConHeaderOfBits = "# of bits";
        private const string ConHeaderOfFuseArray = "# of fuse array";
        private const string ConHeaderProgrammedAt = "Programmed at";
        private const string ConHeaderSupply = "Supply";

        [GeneratedRegex(ConHeaderFuseArray, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(ConHeaderOfBits, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(ConHeaderOfFuseArray, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(ConHeaderProgrammedAt, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex(ConHeaderSupply, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex4();

        private int _fuseArrayColNumber = -1;
        private int _ofBitsColNumber = -1;
        private int _ofFuseArrayColNumber = -1;
        private int _programmedAtColNumber = -1;
        private int _supplyColNumber = -1;

        public override EfuseArraySizeSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition();

            GetHeaders();

            EfuseArraySizeSheet sheet = ReadSheet(sheetName);
            return sheet;
        }

        protected EfuseArraySizeSheet ReadSheet(string sheetName)
        {
            var sheet = new EfuseArraySizeSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new EfuseArraySizeRow
                {
                    RowNum = i,
                    SheetName = sheetName,
                    FuseArray = _fuseArrayColNumber == -1
                        ? ""
                        : ExcelWorksheet.GetCellValue(i, _fuseArrayColNumber).Trim(),
                    OfBits = _ofBitsColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _ofBitsColNumber).Trim(),
                    OfFuseArray = _ofFuseArrayColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _ofFuseArrayColNumber).Trim(),
                    ProgrammedAt = _programmedAtColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _programmedAtColNumber).Trim(),
                    Supply = _supplyColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _supplyColNumber).Trim()
                };

                if (!string.IsNullOrEmpty(row.Supply))
                {
                    row.SupplyPinList.AddRange([.. row.Supply.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries)]);
                }

                sheet.Rows.Add(row);
            }

            sheet.FuseArrayColNumber = _fuseArrayColNumber;
            sheet.OfBitsColNumber = _ofBitsColNumber;
            sheet.OfFuseArrayColNumber = _ofFuseArrayColNumber;
            sheet.ProgrammedAtColNumber = _programmedAtColNumber;
            sheet.SupplyColNumber = _supplyColNumber;
            return sheet;
        }

        protected void GetHeaders()
        {
            for (int i = 1; i <= EndCol; i++)
            {
                string lStrHeader = GetCellValue(StartRow, i).ToUpper().Trim();

                if (MyRegex().IsMatch(lStrHeader))
                {
                    _fuseArrayColNumber = i;
                }
                else if (MyRegex1().IsMatch(lStrHeader))
                {
                    _ofBitsColNumber = i;
                }
                else if (MyRegex2().IsMatch(lStrHeader))
                {
                    _ofFuseArrayColNumber = i;
                }
                else if (MyRegex3().IsMatch(lStrHeader))
                {
                    _programmedAtColNumber = i;
                }
                else if (MyRegex4().IsMatch(lStrHeader))
                {
                    _supplyColNumber = i;
                }
            }
        }

        protected void GetFirstHeaderPosition()
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (MyRegex().IsMatch(GetCellValue(i, j).Trim()))
                    {
                        StartRow = i;
                        break;
                    }
                }
            }
        }

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            object value = ExcelWorksheet.Cells[rowNumber, columnNumber].Value;
            return value?.ToString() ?? "";
        }
    }
}
