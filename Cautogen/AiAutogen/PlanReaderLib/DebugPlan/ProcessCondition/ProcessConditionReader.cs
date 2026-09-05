using System;
using System.Linq;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace DebugPlanReaderLib.DebugPlan
{
    public class ProcessConditionReader : MySheetReader<ProcessConditionSheet>
    {
        private ProcessConditionSheet _sheet;

        public override ProcessConditionSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            string sheetName = excelWorksheet.Name;

            _sheet = new ProcessConditionSheet(sheetName);

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                _sheet.AddDimensionError();
                return _sheet;
            }

            var sheet = new ProcessConditionSheet(sheetName);
            for (int i = StartRow; i <= EndRow; i++)
            {
                string text = ExcelWorksheet.GetCellValue(i, 1).Trim();
                string key = text.Split(':').First();
                if (key.Equals("Efuse Enable word", StringComparison.OrdinalIgnoreCase))
                {
                    sheet.EfuseEnableWord = text.Split(':').Last();
                }
                else if (key.Equals("Tester", StringComparison.OrdinalIgnoreCase))
                {
                    sheet.Tester = text.Split(':').Last();
                }
            }

            return sheet;
        }
    }
}
