using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.Utility;

namespace TestPlanLib.BinCut.NonIgxlSheet
{
    public class BinCutNonIgxlBase(ExcelWorksheet excelWorksheet, string folder = "", bool isCsharp = false)
    {
        protected int ConMaxSearchColumn = 10;
        protected int ConMaxSearchRow = 10;
        protected string SheetName = "";
        protected ExcelWorksheet Worksheet = excelWorksheet;
        protected string StartHeader = "";
        protected string ErrMsg = "";
        protected bool IsCsharp = isCsharp;

        private readonly string _outputFolder = folder;

        public double BaseVoltage { get; set; }

        public double StepSize { get; set; }

        public string WorkFlow(bool isCs, string shadowStage = "")
        {
            string errMsg = "";
            SheetName = Worksheet.Name;

            DataTable table = ReadSheet(Worksheet);
            IsCsharp = isCs;
            EditSheet(table, ref errMsg, IsCsharp);
            if (errMsg?.Length != 0)
            {
                ErrMsg = $"Worksheet {SheetName} {errMsg}";
            }

            string fileName;
            if (SheetName.Contains("post", StringComparison.OrdinalIgnoreCase))
            {
                fileName = ReturnPostFileName(isCs, SheetName);
            }
            else if (SheetName == "EquationVoltages" && IsCsharp)
            {
                fileName = "bincut_ate_condition_eqn_vol";
            }
            else if (SheetName.EqualsIgnoreCase("START_EQN"))
            {
                fileName = "bincut_starting_eqn_table";
            }
            else
            {
                fileName = ReturnFileName(isCs);
            }

            if (!string.IsNullOrEmpty(shadowStage))
            {
                fileName = fileName + "_" + shadowStage;
            }

            BinCutNonIgxlBaseHelpers.WriteToFile(_outputFolder, table, fileName, IsCsharp);

            return Path.Combine(_outputFolder, fileName + ".txt");
        }

        protected virtual void EditSheet(DataTable dataTable, ref string errMsg, bool isCsharp)
        {
            RemoveBlankTail(dataTable);
        }

        protected virtual string ReturnFileName(bool isCs)
        {
            return SheetName;
        }

        protected virtual string ReturnPostFileName(bool isCs, string sheetName)
        {
            return SheetName;
        }

        protected virtual DataTable ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var dt = new DataTable
            {
                TableName = excelWorksheet.Name
            };
            ExcelCellAddress startCell = excelWorksheet.Dimension.Start;
            ExcelCellAddress endCell = excelWorksheet.Dimension.End;

            for (int col = 1; col <= endCell.Column; col++)
            {
                dt.Columns.Add(col.ToString(CultureInfo.InvariantCulture));
            }

            for (int row = 1; row <= endCell.Row; row++)
            {
                DataRow dr = dt.NewRow();
                int x = 0;
                for (int col = startCell.Column; col <= endCell.Column; col++)
                {
                    object cellValue = excelWorksheet.Cells[row, col].Value;
                    if (cellValue != null)
                    {
                        cellValue = excelWorksheet.GetCellValue(row, col);
                    }

                    dr[x++] = cellValue;
                }
                dt.Rows.Add(dr);
            }

            ConMaxSearchRow = ConMaxSearchRow > dt.Rows.Count ? dt.Rows.Count : ConMaxSearchRow;
            ConMaxSearchColumn = ConMaxSearchColumn > dt.Columns.Count ? dt.Columns.Count : ConMaxSearchColumn;

            return dt;
        }

        protected static void RemoveBlankTail(DataTable dataTable)
        {
            for (int i = dataTable.Rows.Count - 1; i > 0; i--)
            {
                DataRow row = dataTable.Rows[i];
                if (row.ItemArray.Any(x => !string.IsNullOrEmpty(x?.ToString())))
                {
                    break;
                }
                else
                {
                    dataTable.Rows.RemoveAt(i);
                }
            }
        }
    }
}
