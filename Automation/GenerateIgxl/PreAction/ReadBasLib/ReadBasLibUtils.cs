using System.Collections.Generic;
using System.Linq;

using Automation.Static;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.PreAction.ReadBasLib
{
    public static class ReadBasLibUtils
    {
        public static void InitialParamMapping()
        {
            if (SettingStatic.BasicConfigWorkbook == null)
            {
                return;
            }
            ExcelWorksheet wSheet = SettingStatic.BasicConfigWorkbook.Worksheets[NeededSheets.ParamMap];
            if (wSheet == null)
            {
                return;
            }

            const int standardCol = 1;
            const int aliasCol = 2;
            for (int i = 2; i <= wSheet.Dimension.End.Row; i++)
            {
                if (wSheet.Cells[i, standardCol].Value == null || wSheet.Cells[i, aliasCol].Value == null)
                {
                    continue;
                }

                string standardName = wSheet.Cells[i, standardCol].Value.ToString();
                string aliasName = wSheet.Cells[i, aliasCol].Value.ToString();
                var paraNameGroup = new List<string> { standardName };
                paraNameGroup.AddRange(aliasName.Split(',').ToList());
                VbtFunctionLibShared.ParamMappingList.Add(paraNameGroup);
            }
        }
    }
}
