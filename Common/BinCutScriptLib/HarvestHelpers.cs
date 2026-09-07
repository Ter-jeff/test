using System;
using System.Drawing;

using BinCutScriptLib.Static;

using IgxlLib.Utility;

using OfficeOpenXml;

using TestPlanLib;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;

namespace BinCutScriptLib
{
    internal class HarvestHelpers
    {
        internal static void ReadHarvestDssc(string testProgramFilePath, Job job, Action<string, Color> richTextBoxAppend)
        {
            ExcelWorksheet? harvMappingTableSheet =
                IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "HARVMappingTable")
                ?? IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "HARVMappingTable_" + job.JobName);

            if (harvMappingTableSheet != null)
            {
                richTextBoxAppend("Reading HARVMappingTable ...", Color.Blue);
                BinCutData.HarvMappingTableSheet = new HarvMappingTableReader().ReadSheet(harvMappingTableSheet);
            }

            ExcelWorksheet? harvPmodeTableSheet =
                IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "HARV_Pmode_Table")
                ?? IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "HARV_Pmode_Table_" + job.JobName);

            if (harvPmodeTableSheet != null)
            {
                richTextBoxAppend("Reading HARV_Pmode_Table ...", Color.Blue);
                BinCutData.HarvPmodeTableFstpSheet = new HarvPmodeTableFstpReader().ReadSheet(harvPmodeTableSheet);
                BinCutData.HarvPmodeTableHarvCheckSheet =
                    new HarvPmodeTableHarvCheckReader().ReadSheet(harvPmodeTableSheet);
                BinCutData.HarvPmodeTableGroupSheet = new HarvPmodeTableGroupReader(
                    BinCutData.HarvPmodeTableHarvCheckSheet!.HarvCount,
                    BinCutData.HarvMappingTableSheet!.BincutPatternCount
                ).ReadSheet(harvPmodeTableSheet);
                //BinCutData.HarvPmodeTableModeSheet = new HarvPmodeTableModeReader().ReadSheet(
                //    harvPmodeTableSheet
                //);
                BinCutData.HarvPmodeTableOffsetSheet = new HarvPmodeTableOffsetReader().ReadSheet(harvPmodeTableSheet);
            }
        }
    }
}
