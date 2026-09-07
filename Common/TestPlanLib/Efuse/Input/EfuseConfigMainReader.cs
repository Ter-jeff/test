using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Efuse.Input
{
    public partial class EfuseConfigMainReader : MySheetReader<EfuseConfigMainSheet>
    {
        private const string ConHeaderFuse = "^Fuse$";
        private const string ConHeaderDescription = "^Description$";
        private const string ConHeaderValueToFuse = "^value_to_fuse$";
        private const string ConHeaderValueToFuse2 = "1st.*Silicon";
        private const string ConHeaderFuseBlowLocation = @"^Fuse Blow.*\n*Location$";

        [GeneratedRegex("^x", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(ConHeaderFuse, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(ConHeaderDescription, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(ConHeaderValueToFuse, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex(ConHeaderValueToFuse2, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex4();
        [GeneratedRegex(ConHeaderFuseBlowLocation, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex5();

        private int _fuseColNumber = -1;
        private int _descriptionColNumber = -1;
        private int _a00StartColNumber;
        private int _a00ColCount;
        private int _valueToFuseColNumber = -1;
        private int _fuseBlowLocationColNumber = -1;

        public override EfuseConfigMainSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition();

            GetHeaders();

            EfuseConfigMainSheet sheet = ReadSheet(sheetName);
            return sheet;
        }

        protected EfuseConfigMainSheet ReadSheet(string sheetName)
        {
            var sheet = new EfuseConfigMainSheet(sheetName) { AccessMode = GetCellValue(1, 1).Trim() };

            //read header
            var header = new List<string>();
            for (int iCol = 1; iCol <= EndCol; iCol++)
            {
                object value = ExcelWorksheet.Cells[StartRow, iCol].Value ?? "";
                header.Add(value.ToString() ?? "");
            }
            sheet.Header = string.Join("\t", header);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new EfuseConfigMainRow { RowNum = i, SheetName = sheetName, Fuse = _fuseColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _fuseColNumber).Trim() };
                if (!string.IsNullOrEmpty(row.Fuse))
                {
                    var lsbmsb = new EFuseLsbMsb();
                    lsbmsb.SetLsbmsbData(row.Fuse);
                    row.Lsb = int.Parse(lsbmsb.GetLsb());
                    row.Msb = int.Parse(lsbmsb.GetMsb());
                    row.Width = row.Msb - row.Lsb + 1;
                    row.NeedSort = lsbmsb.GetSortFlg();
                }
                row.Description = _descriptionColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _descriptionColNumber).Trim();

                row.ValueToFuse = _valueToFuseColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _valueToFuseColNumber).Trim();
                row.FuseBlowLocation = _fuseBlowLocationColNumber == -1 ? "" : ExcelWorksheet.GetCellValue(i, _fuseBlowLocationColNumber).Trim();
                if (row.ValueToFuse.ContainsIgnoreCase("SVM"))
                {
                    sheet.HasSvm = true;
                }

                if (_a00StartColNumber != 0)
                {
                    for (int col = 0; col < _a00ColCount; col++)
                    {
                        string data = ExcelWorksheet.GetCellValue(i, _a00StartColNumber + col).Trim();
                        if (data.StartsWithIgnoreCase("x"))
                        {
                            data = MyRegex().Replace(data, "0x");
                        }
                        else if (data.StartsWithIgnoreCase("b"))
                        {
                            data = data.EFuseBitToInt();
                        }

                        row.DataList.Add(data);
                    }
                }
                row.IsAllSameData = IsAllSameData(row);
                sheet.Rows.Add(row);
            }

            sheet.FuseColNumber = _fuseColNumber;
            sheet.DescriptionColNumber = _descriptionColNumber;
            sheet.DataColCount = _a00ColCount;
            sheet.DataStartCol = _a00StartColNumber;

            sheet.ValueToFuseColNumber = _valueToFuseColNumber;
            sheet.FuseBlowLocationColNumber = _fuseBlowLocationColNumber;
            return sheet;
        }

        public static bool IsAllSameData(EfuseConfigMainRow efuseConfigMainRow)
        {
            return efuseConfigMainRow.DataList.All(x => x.EqualsIgnoreCase(efuseConfigMainRow.DataList.First()));
        }

        protected void GetHeaders()
        {
            for (int i = 1; i <= EndCol; i++)
            {
                string lStrHeader = GetCellValue(StartRow, i).ToUpper().Trim();

                if (MyRegex1().IsMatch(lStrHeader))
                {
                    _fuseColNumber = i;
                }
                else if (MyRegex2().IsMatch(lStrHeader))
                {
                    _descriptionColNumber = i;
                }
                else if (MyRegex3().IsMatch(lStrHeader) || MyRegex4().IsMatch(lStrHeader))
                {
                    _valueToFuseColNumber = i;
                }
                else if (MyRegex5().IsMatch(lStrHeader))
                {
                    _fuseBlowLocationColNumber = i;
                }
                //else if (Regex.IsMatch(lStrHeader, ConHeaderA00, RegexOptions.IgnoreCase))
                else if (_descriptionColNumber != -1 && _valueToFuseColNumber == -1)
                {
                    if (_a00StartColNumber == 0)
                    {
                        _a00StartColNumber = i;
                    }

                    _a00ColCount++;
                }
            }
        }

        protected void GetFirstHeaderPosition()
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (MyRegex1().IsMatch(GetCellValue(i, j).Trim()))
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
