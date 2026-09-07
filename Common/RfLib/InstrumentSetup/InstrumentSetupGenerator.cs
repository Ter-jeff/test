using System.Collections.Generic;
using System.Data;

using CommonLib.Extension;

using OfficeOpenXml;

namespace RfLib.InstrumentSetup
{
    internal class InstrumentSetupGenerator(ExcelPackage excelPackage)
    {
        public ExcelPackage XlPackage = excelPackage;
        public ExcelWorkbook XlWorkBook = excelPackage.Workbook;

        public void GenInstrumentSetupTable(DataTable dataTable)
        {
            const int instrumentTypeIndex = 4;
            ExcelWorksheet ws = XlPackage.Workbook.Worksheets["InstrumentSetup"] ??
                                       XlPackage.Workbook.Worksheets.Add("InstrumentSetup");
            ws.Cells[1, 1].LoadFromDataTable(dataTable, false);

            ExcelWorksheet instruementTable = XlPackage.Workbook.Worksheets["InstrumentSetup"];

            if (instruementTable.Dimension == null)
            {
                return;
            }

            for (int i = instrumentTypeIndex; i < instruementTable.Dimension.Rows; i++)
            {
                List<string> typeList = [.. (instruementTable.Cells[i, instrumentTypeIndex].Value?.ToString() ?? "").Split(',')];
                if (typeList.Count > 1)
                {
                    OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList instType = instruementTable.Cells[i, instrumentTypeIndex].DataValidation.AddListDataValidation();
                    foreach (string type in typeList)
                    {
                        instType.Formula.Values.Add(type);
                    }
                    instruementTable.Cells[i, instrumentTypeIndex].Value = "";
                }
            }
            ws.Cells.TryAutoFitColumns();
        }
    }
}
