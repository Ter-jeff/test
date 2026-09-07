using Automation.InputManager.Data;
using Automation.Static;

using OfficeOpenXml;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.CPM;
using TestPlanLib.Scan;
using TestPlanLib.Static;

namespace Automation.InputManager
{
    public class ScanInputManager : InputManagerBase<ScanInputData>
    {
        public ScanInputManager(ExcelWorkbook excelWorkbook) : base(excelWorkbook)
        {
        }

        public override ScanInputData Read()
        {
            var result = new ScanInputData
            {
                ScanConfig = SettingStatic.ScanConfig
            };
            if (EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.ClockPllMeas] != null)
            {
                var clockOutReader = new ClockMeasReader();
                result.ClockMeasSheet = clockOutReader.ReadSheet(EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.ClockPllMeas]);
            }
            if (EpWorkbook.TestPlanWorkbook.Worksheets["Instance_CPM"] != null && LocalSpecs.IsModuleIncluded("CPM"))
            {
                if (EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_cpm"] != null)
                {
                    var efuseCpmReader = new EfuseCpmSheetReader();
                    result.EfuseCpmSheet = efuseCpmReader.ReadSheet(EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_cpm"]);

                }
            }
            if (EpWorkbook.TestPlanWorkbook.Worksheets["TurboModeCheck"] != null)
            {
                var turboModeSheetReader = new BinCutInstanceSheetReader();
                result.TurboModeInstanceSheet = turboModeSheetReader.ReadSheet(EpWorkbook.TestPlanWorkbook.Worksheets["TurboModeCheck"]);
            }

            return result;
        }
    }
}
