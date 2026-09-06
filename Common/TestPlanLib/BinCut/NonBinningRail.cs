using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

using OfficeOpenXml;

using TestPlanLib.BinCut.NonIgxlSheet;

namespace TestPlanLib.BinCut
{
    public partial class NonBinningRail : BinCutNonIgxlBase
    {
        #region Field
        private const string ConHeaderPerfMode = "Performance.*";

        [GeneratedRegex(ConHeaderPerfMode, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();

        private int _headerRow;
        private int _firstModeColumn;
        #endregion

        public NonBinningRail(string pFolder, ExcelWorksheet excelWorksheet)
            : base(excelWorksheet, pFolder)
        {
            StartHeader = ConHeaderPerfMode;
        }

        protected override void EditSheet(DataTable dataTable, ref string errMsg, bool isCsharp)
        {
            RemoveBlankTail(dataTable);
            var perforModeList = new List<string>();

            FindFirstModeHeader(dataTable);

            ReadModeList(perforModeList, dataTable);

            //for (int i = _firstModeColumn + 1; i < pTable.Columns.Count; i++)
            //{
            //    if (Regex.IsMatch(pTable.Rows[_headerRow][i].ToString(), ConHeaderPerfMode, RegexOptions.IgnoreCase))
            //    {
            //        FillBlankCell(perforModeList, i, pTable);
            //    }
            //}
        }

        protected override string ReturnFileName(bool isCs)
        {
            return !isCs ? BinCutConst.ConNonBinningRailFileName : BinCutConst.ConBincutAteConditionFileName;
        }

        private void FindFirstModeHeader(DataTable dataTable)
        {
            for (int i = 0; i < ConMaxSearchRow; i++)
            {
                for (int j = 0; j < ConMaxSearchColumn; j++)
                {
                    if (MyRegex().IsMatch(dataTable.Rows[i][j].ToString()!))
                    {
                        _headerRow = i;
                        _firstModeColumn = j;
                        return;
                    }
                }
            }
        }

        private void ReadModeList(List<string> modeList, DataTable dataTable)
        {
            for (int i = _headerRow + 1; i < dataTable.Rows.Count; i++)
            {
                modeList.Add(dataTable.Rows[i][_firstModeColumn].ToString()!);
            }
        }
    }
}
