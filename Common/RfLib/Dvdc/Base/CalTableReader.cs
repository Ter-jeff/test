using System.Collections.Generic;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace RfLib.Dvdc.Base
{
    public partial class CalTableReader
    {
        [GeneratedRegex("Path Name", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("LUT", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        private int _colPathIndex = -1;
        private int _colLutIndex = -1;
        public Dictionary<string, string> CalPathToLut = [];

        public void Read(ExcelWorksheet excelWorksheet)
        {
            int rowindex = 1;
            ReadHeader(excelWorksheet, ref rowindex);
            ReadContent(excelWorksheet, rowindex);

        }

        private void ReadHeader(ExcelWorksheet excelWorksheet, ref int rowIndex)
        {
            //Path names	LUT
            for (rowIndex = 1; rowIndex <= excelWorksheet.Dimension.Rows; rowIndex++)
            {
                if (_colLutIndex != -1 && _colPathIndex != -1)
                {
                    break;
                }

                for (int j = 1; j <= excelWorksheet.Dimension.Columns; j++)
                {
                    if (string.IsNullOrEmpty(excelWorksheet.Cells[rowIndex, j].Text))
                    {
                        continue;
                    }

                    if (MyRegex().IsMatch(excelWorksheet.Cells[rowIndex, j].Text))
                    {
                        _colPathIndex = j;
                    }

                    if (MyRegex1().IsMatch(excelWorksheet.Cells[rowIndex, j].Text))
                    {
                        _colLutIndex = j;
                    }
                }

            }
        }

        private void ReadContent(ExcelWorksheet excelWorksheet, int rowIndex)
        {
            for (int i = rowIndex; i <= excelWorksheet.Dimension.Rows; i++)
            {
                //var path = pathRename(sheet.Cells[i, _colPathIndex].Text).Replace("-","_to_");
                string path = excelWorksheet.Cells[i, _colPathIndex].Text.Replace("-", "_to_");
                string lut = excelWorksheet.Cells[i, _colLutIndex].Text;
                CalPathToLut.TryAdd(path, lut);
            }
        }
    }
}
