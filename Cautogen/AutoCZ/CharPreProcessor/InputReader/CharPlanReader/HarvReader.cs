using System.Collections.Generic;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader
{
    public class HarvReader
    {
        public static List<string> HarvItems = new List<string>();

        public void ReadMappingTable(ExcelWorksheet sheet)
        {
            HarvItems.Clear();

            // read header
            int startCol = _SearchTableStartRow(sheet);
            int endCol = sheet.Dimension.End.Column;

            // early return
            if (startCol == endCol)
            {
                return;
            }

            _ReadTableHeader(sheet, startCol);
        }

        private void _ReadTableHeader(ExcelWorksheet sheet, int startcol)
        {
            for (int i = startcol + 1; i <= sheet.Dimension.End.Column; i++)
            {
                if (sheet.Cells[1, i].Value != null)
                {
                    HarvItems.Add(sheet.Cells[1, i].Value.ToString().Trim());
                }
            }
        }

        private static int _SearchTableStartRow(ExcelWorksheet sheet)
        {
            /* Return the row number of Column A == "Block" */
            for (int i = 1; i <= sheet.Dimension.End.Column; i++)
            {
                if (sheet.Cells[1, i].Value != null && sheet.Cells[1, i].Value.ToString().Trim() == "Harvest Result0")
                {
                    return i;
                }
            }
            return sheet.Dimension.End.Column;
        }

    }
}
