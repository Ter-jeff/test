using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;
using Automation.Reader.ConfigFile.RtosTable;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.Business.LevelRows
{
    public class RtosLevelsGenerator : LevelsGenerator
    {
        public override void UpdateLevelSheet(ref LevelSheet levelSheet, LevelData levelData = null)
        {
            ExcelWorksheet rtosTableSheet = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Find(x => x.Name.Equals("Rtos_Table", StringComparison.OrdinalIgnoreCase));
            if (rtosTableSheet == null)
            {
                return;
            }

            List<RtosLevelRow> rtosTableLevels = RtosTableSheet.LoadConfig(rtosTableSheet).LevelRows;
            foreach (RtosLevelRow row in rtosTableLevels)
            {
                string pinName = row.PinName;

                foreach (KeyValuePair<string, string> param in row.Parameters)
                {
                    levelSheet.Rows.Add(new LevelRow(pinName, param.Key, param.Value, "Rtos_Table"));
                }
            }
        }
    }
}
