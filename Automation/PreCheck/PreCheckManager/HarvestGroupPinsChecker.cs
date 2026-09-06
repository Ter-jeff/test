using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Atpg;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;

namespace Automation.PreCheck.PreCheckManager
{
    public class HarvestGroupPinsChecker
    {
        public void Check()
        {
            List<string> harvestCheckedPattern = new List<string>();
            Dictionary<string, PatternData> patternList = AcTSetCategoryMapSingleton.Instance().PatternList;
            Dictionary<string, int> timeSets = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
            var ufDigSrcSheet = new List<UfDigSrcSheet>();
            if (TestPlanStatic.UfDigSrcSheets != null && TestPlanStatic.UfDigSrcSheets.Any())
            {
                ufDigSrcSheet = TestPlanStatic.UfDigSrcSheets;
            }

            ExcelPackage missingPinsReport = InitMissPinReport();
            foreach (BinCutInstanceSheet binCutInstanceSheet in TestPlanStatic.ScanInstanceSheets)
            {
                new NonBinCutInstanceSheetChecker().WorkFlow(binCutInstanceSheet, patternList, timeSets, ufDigSrcSheet);
                if (binCutInstanceSheet != null)
                {
                    binCutInstanceSheet.AddToErrorReport();
                    if (binCutInstanceSheet.HasPatternPinGroups())
                    {
                        AtpgService.CheckHarvestGroupPins(
                            binCutInstanceSheet,
                            missingPinsReport,
                            harvestCheckedPattern);
                    }
                }
            }
            if (harvestCheckedPattern.Any())
            {
                missingPinsReport.Save();
            }
        }

        private ExcelPackage InitMissPinReport()
        {
            string outputPath = Path.Combine(LocalSpecs.TarFolder, "MissingPinReport.xlsx");
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            var missingPinsReport = new ExcelPackage(new FileInfo(outputPath));
            ExcelWorksheet extraSheet = missingPinsReport.Workbook.AddSheet("Extra harvest pins");
            ExcelWorksheet missSheet = missingPinsReport.Workbook.AddSheet("Miss harvest pins");
            var titles = new List<string> { "Pattern", "Pins" };
            extraSheet.Cells[1, 1].PrintExcelRow(titles.ToArray());
            missSheet.Cells[1, 1].PrintExcelRow(titles.ToArray());
            return missingPinsReport;
        }
    }
}
