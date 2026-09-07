using System.Collections.Generic;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.DataStruct
{
    public class IoIgnoreList
    {
        //"Pin Name [Regular Expression]" "Ignore type [Bypass\\Warning]"
        public Dictionary<string, string>? IoIgnorePinsTypes;

        //"Pin Name [Regular Expression]" "Ignore sheet [Sheet Name]"
        public Dictionary<string, string>? IoIgnorePinsSheets;

        public IoIgnoreList(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null || excelWorksheet.Dimension == null)
            {
                IoIgnorePinsTypes = null;
                IoIgnorePinsSheets = null;
                return;
            }

            IoIgnorePinsTypes = new Dictionary<string, string>() { { "", "Error" } };
            IoIgnorePinsSheets = [];

            for (int i = excelWorksheet.Dimension.Start.Row + 1; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                string pinname = EpplusExtensions.GetMergedCellValue(excelWorksheet, i, 1).Trim();
                string ignoretype = EpplusExtensions.GetMergedCellValue(excelWorksheet, i, 2).Trim();
                string ignoresheet = EpplusExtensions.GetMergedCellValue(excelWorksheet, i, 3).Trim();

                if (string.IsNullOrWhiteSpace(pinname))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(ignoretype))
                {
                    if (!IoIgnorePinsTypes.TryAdd(pinname, ignoretype))
                    {
                        IoIgnorePinsTypes[pinname] = ignoretype;
                    }
                }

                if (!string.IsNullOrWhiteSpace(ignoresheet))
                {
                    if (!IoIgnorePinsSheets.TryAdd(pinname, ignoresheet))
                    {
                        IoIgnorePinsSheets[pinname] = ignoresheet;
                    }
                }

            }
        }
    }
}
