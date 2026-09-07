using System;
using System.Linq;

using Automation.Static;

using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.Singleton
{
    public class ExceptionListSingleton
    {
        internal static ExceptionListSingleton _instance;
        private const string ConDataLocation = "ExceptionList";
        private const string ConSheetName = "InstanceExceptionList";
        private readonly InstanceExceptionList _exceptionList;

        private ExceptionListSingleton()
        {
            _exceptionList = new InstanceExceptionList();
        }

        public static ExceptionListSingleton Instance()
        {
            if (_instance == null)
            {
                lock (typeof(ExceptionListSingleton))
                {
                    if (_instance == null)
                    {
                        _instance = new ExceptionListSingleton();
                    }
                }
            }
            return _instance;
        }

        public void Initialize()
        {
            if (EpWorkbook.TestPlanWorkbook == null || !EpWorkbook.TestPlanWorkbook.Worksheets.Any(p => p.Name.Equals(ConSheetName)))
            {
                return;
            }

            _exceptionList.ExceptionList.Clear();
            ExcelWorksheet exceptionSheet = EpWorkbook.TestPlanWorkbook.Worksheets[ConSheetName];
            if (exceptionSheet == null)
            {
                return;
            }

            int startRow = 1, startColumn = 1;
            FindCellByValue(ref startRow, ref startColumn, exceptionSheet, ConDataLocation);
            for (int i = startRow + 2; i <= exceptionSheet.Dimension.Rows; i++)
            {
                InstanceExceptionEntry entry = new InstanceExceptionEntry { TestName = EpplusExtensions.GetCellValue(exceptionSheet, i, startColumn) };
                if (entry.TestName == string.Empty)
                {
                    break; //terminate when meet blank row
                }

                entry.DcCategory = EpplusExtensions.GetCellValue(exceptionSheet, i, startColumn + 1);
                _exceptionList.AddRow(entry);
            }
        }

        public static void FindCellByValue(ref int row, ref int col, ExcelWorksheet sheet, string value)
        {
            for (int i = 1; i <= sheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= sheet.Dimension.End.Column; j++)
                {
                    string cell = sheet.GetCellValue(i, j);
                    if (cell.Equals(value, StringComparison.OrdinalIgnoreCase))
                    {
                        row = i;
                        col = j;
                        return;
                    }
                }
            }
        }

        public string GetDcCategoryByInstance(string insName)
        {
            InstanceExceptionEntry entry = _exceptionList.ExceptionList.Find(x => x.TestName.Equals(insName, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                return string.Empty;
            }

            return entry.DcCategory;
        }
    }
}
