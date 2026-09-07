using System.Collections.Generic;

using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.InputManager
{
    public class HtolInputManager : InputManagerBase<HtolInputData>
    {
        public HtolInputManager(ExcelWorkbook excelWorkbook) : base(excelWorkbook)
        {
        }

        public override HtolInputData Read()
        {
            var result = new HtolInputData();
            var binCutInstanceSheets = new List<BinCutInstanceSheet>();
            Dictionary<string, PatternData> patternList = AcTSetCategoryMapSingleton.Instance().PatternList;
            Dictionary<string, int> timeSets = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
            foreach (BinCutInstanceSheet sheet in TestPlanStatic.HtolInstanceSheets)
            {
                new NonBinCutInstanceSheetChecker().WorkFlow(sheet, patternList, timeSets);
                if (sheet != null)
                {
                    binCutInstanceSheets.Add(sheet);
                    sheet.AddToErrorReport();
                }
            }
            result.BinCutInstanceSheets = binCutInstanceSheets;

            return result;
        }
    }
}
