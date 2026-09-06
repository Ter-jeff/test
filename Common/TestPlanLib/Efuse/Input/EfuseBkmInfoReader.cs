using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.Efuse.Input
{
    public class EfuseBkmInfoReader
    {
        private int _bkmcolumnIndex = -1;
        private int _startRowindex;
        public ExcelWorksheet? BkmSheet;

        public EfuseBkmInfoReader(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return;
            }

            BkmSheet = excelWorksheet;
            ReadHeader();
        }

        private void ReadHeader()
        {
            for (_startRowindex = 1; _startRowindex <= BkmSheet!.Dimension.Rows; _startRowindex++)
            {
                for (int col = 1; col <= BkmSheet.Dimension.Columns; col++)
                {
                    if (string.IsNullOrEmpty(BkmSheet.Cells[_startRowindex, col].Text))
                    {
                        continue;
                    }

                    if (BkmSheet.Cells[_startRowindex, col].Text.Trim()
.EqualsIgnoreCase("bkm"))
                    {
                        _bkmcolumnIndex = col;
                    }
                }
                if (_bkmcolumnIndex != -1)
                {
                    break;
                }
            }
            _startRowindex++;
        }

        public void Check()
        {
            if (BkmSheet == null || _bkmcolumnIndex == -1)
            {
                return;
            }

            string regBkm = @"(?<main>BKM\d+)(?<portion>\.\d)?";
            for (int row = _startRowindex; row <= BkmSheet.Dimension.Rows; row++)
            {
                if (string.IsNullOrEmpty(BkmSheet.Cells[row, _bkmcolumnIndex].Text))
                {
                    continue;
                }

                string value = BkmSheet.Cells[row, _bkmcolumnIndex].Text.Replace(" ", "");
                var match = new Regex(regBkm, RegexOptions.IgnoreCase);
                if (!string.IsNullOrEmpty(match.Match(value).Groups["main"].Value))
                {
                    if (string.IsNullOrEmpty(match.Match(value).Groups["portion"].Value))
                    {
                        BkmSheet.Cells[row, _bkmcolumnIndex].Value = value + ".0";
                        ErrorReportManager.AddError(EFuseErrorType.E_InvalidValue_01, BkmSheet.Name, row, row, [value, BkmSheet.Cells[row, _bkmcolumnIndex].Value?.ToString() ?? ""]);
                    }
                }
            }
        }

        public void Generate(string dir, string filename)
        {
            if (BkmSheet == null)
            {
                return;
            }

            string path = Path.Combine(dir, filename + ".txt");
            int maxRow = BkmSheet.Dimension.Rows;
            int maxCol = BkmSheet.Dimension.Columns;
            StreamWriter sw = File.CreateText(path);
            for (int iRow = 1; iRow <= maxRow; iRow++)
            {
                var row = new List<string>();
                for (int iCol = 1; iCol <= maxCol; iCol++)
                {
                    object value = BkmSheet.Cells[iRow, iCol].Value ?? "";
                    row.Add(value.ToString() ?? "");
                }
                sw.WriteLine(string.Join("\t", row));
            }
            sw.Close();
        }
    }
}
