using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PreAction.GenDSSCSetup;
using Automation.Static;

using CommonLib.Extension;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib.BinCut.NonIgxlSheet;

namespace Automation.GenerateIgxl.PostAction.GenNonIgxlSheets
{
    internal class GenNonIgxlSheetsMain
    {
        public void WorkFlow()
        {
            if (LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                EpWorkbook.EquationVoltages = new ExcelPackage(new FileInfo(LocalSpecs.EquationVoltagesFileName)).Workbook;
                ExcelWorksheet worksheet = EpWorkbook.EquationVoltages.Worksheets["EquationVoltages"];
                if (worksheet != null)
                {
                    var nonIgxlSheet = new BinCutNonIgxlBase(worksheet, FolderStructure.DirCommon, !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder));
                    bool isCs = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder);
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, Path.GetFileNameWithoutExtension(nonIgxlSheet.WorkFlow(isCs)));
                }
            }

            GenerateRelatedFiles();
            GenerateTimeSettingsSheet();

            if (TestPlanStatic.DigitalFlagsSheet != null)
            {
                TestPlanStatic.DigitalFlagsSheet.ExpandCoreRow();
                TestPlanStatic.DigitalFlagsSheet.ExportToTxt(Path.Combine(FolderStructure.DirCommonSheets, "Digital_Flags.txt"));
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, "Digital_Flags");
            }

            // Add DsscSetup sheets
            Dictionary<string, List<DsscSetupSheet>> genDsscSetupSheets = DsscSetupMain.GetSheets();
            if (genDsscSetupSheets != null && genDsscSetupSheets.Any())
            {
                List<string> dsscsheets = DsscSetupMain.ExportAllSheets(FolderStructure.DirHardIp);
                foreach (string file in dsscsheets)
                {
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirHardIp, file);
                }
            }

            if (SettingStatic.BasicConfigWorkbook.Worksheets["EnableWdGatingTable"] != null)
            {
                ExcelWorksheet enableWdGatingTb = SettingStatic.BasicConfigWorkbook.Worksheets["EnableWdGatingTable"];
                if (enableWdGatingTb != null)
                {
                    enableWdGatingTb.ExportWorkBook2Txt(FolderStructure.DirCommon);
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, enableWdGatingTb.Name);
                }
            }
        }
        private void GenerateRelatedFiles()
        {
            ExportUniqueTable("IDS_Mapping_Table");
            ExportUniqueTable("BinOutCalcScanMbistTable");
            ExportTable(EpWorkbook.TestPlanWorkbook.Worksheets["HarvestPinFlag_Table"]);
            ExportChecklistSheets();
            ExportDramTypeSheet();
            ExportFuseCheckSheet();
            ExportTable(EpWorkbook.TestPlanWorkbook.Worksheets["HarvestPinGrpFlagTable"]);
            ExportTablesByPrefix("HARVMappingTable");
            ExportTablesByPrefix("HARV_PMode_Table");
            ExportTable(EpWorkbook.TestPlanWorkbook.Worksheets["Dig_Src_instructions"]);
            ExportUfDigSrcSheets();
            ExportTable(EpWorkbook.TestPlanWorkbook.Worksheets["external_files_sync_manifest"]);
        }

        private void ExportTable(ExcelWorksheet sheet)
        {
            if (sheet != null)
            {
                sheet.ExportWorkBook2Txt(FolderStructure.DirCommon);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, sheet.Name);
            }
        }

        private void ExportUniqueTable(string sheetName)
        {
            ExcelWorksheet table = EpWorkbook.TestPlanWorkbook.Worksheets[sheetName];
            if (table != null && !TestProgram.NonIgxlSheetsList.SheetList.Exists(x => x == Path.Combine(FolderStructure.DirCommon, sheetName)))
            {
                table.ExportWorkBook2Txt(FolderStructure.DirCommon);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, table.Name);
            }
        }

        private void ExportChecklistSheets()
        {
            IEnumerable<string> checkListSheets = EpWorkbook.TestPlanWorkbook.Worksheets.Where(x => x.Name.EndsWith("_Flow_ChkList", StringComparison.OrdinalIgnoreCase)).Select(x => x.Name);
            foreach (string checkListSheet in checkListSheets)
            {
                ExportUniqueTable(checkListSheet);
            }
        }

        private void ExportDramTypeSheet()
        {
            if (!string.IsNullOrEmpty(LocalSpecs.DramTypeFileName) && LocalSpecs.DramTypeFileName != "N/A")
            {
                var dramTypeExcel = new ExcelPackage(new FileInfo(LocalSpecs.DramTypeFileName));
                ExcelWorksheet sheet = dramTypeExcel.Workbook.Worksheets["DRAM_CONFIG"] ?? dramTypeExcel.Workbook.Worksheets["DRAM_Table"];
                if (sheet != null)
                {
                    sheet.Name = "DRAM_Table";
                    sheet.ExportWorkBook2Txt(FolderStructure.DirCommon);
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, sheet.Name);
                }
            }
        }

        private void ExportFuseCheckSheet()
        {
            if (!string.IsNullOrEmpty(LocalSpecs.FuseCheckFileName) && LocalSpecs.FuseCheckFileName != "N/A")
            {
                var fuseCheckExcel = new ExcelPackage(new FileInfo(LocalSpecs.FuseCheckFileName));
                ExportTable(fuseCheckExcel.Workbook.Worksheets["FuseCheckTable"]);
            }
        }

        private void ExportTablesByPrefix(string prefix)
        {
            List<ExcelWorksheet> tables = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().FindAll(x => x.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
            foreach (ExcelWorksheet table in tables)
            {
                table.ExportWorkBook2Txt(FolderStructure.DirCommon);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, table.Name);
            }
        }

        private void ExportUfDigSrcSheets()
        {
            List<ExcelWorksheet> ufDigSrc = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().FindAll(x => x.Name.StartsWith("UF_DigSrc_", StringComparison.OrdinalIgnoreCase));
            if (ufDigSrc.Count <= 1)
            {
                if (ufDigSrc.Count == 1)
                {
                    ufDigSrc.First().ExportWorkBook2Txt(FolderStructure.DirCommon);
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, ufDigSrc.First().Name);
                }
            }
            else
            {
                ErrorMessageBox.Show("Test program can only has one UF_DigSrc sheet but there're multiple in test plan, please check after generation complete!!!");
                foreach (ExcelWorksheet sheet in ufDigSrc)
                {
                    sheet.ExportWorkBook2Txt(FolderStructure.DirCommon);
                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, sheet.Name);
                }
            }
        }

        private void GenerateTimeSettingsSheet()
        {
            ExcelWorksheet timeSettings = EpWorkbook.TestPlanWorkbook.Worksheets["TimeSettings"] ?? EpWorkbook.TestPlanWorkbook.Worksheets["TimeSettings"];
            if (timeSettings != null)
            {
                timeSettings.ExportWorkBook2Txt(FolderStructure.DirCommon);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, timeSettings.Name);
            }
        }

    }
}
