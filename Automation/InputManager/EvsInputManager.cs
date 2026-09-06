using System.Collections.Generic;

using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.InputManager
{
    public class EvsInputManager : InputManagerBase<EvsInputData>
    {
        public EvsInputManager(ExcelWorkbook excelWorkbook) : base(excelWorkbook)
        {
        }

        public override EvsInputData Read()
        {
            var result = new EvsInputData();
            var evsInstanceSheets = new List<BinCutInstanceSheet>();
            Dictionary<string, PatternData> patternList = AcTSetCategoryMapSingleton.Instance().PatternList;
            Dictionary<string, int> timeSets = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
            foreach (BinCutInstanceSheet evsInstanceSheet in TestPlanStatic.EvsInstanceSheets)
            {
                new NonBinCutInstanceSheetChecker().WorkFlow(evsInstanceSheet, patternList, timeSets);
                if (evsInstanceSheet != null)
                {
                    evsInstanceSheets.Add(evsInstanceSheet);
                    evsInstanceSheet.AddToErrorReport();
                }
            }
            result.EvsInstanceSheets = evsInstanceSheets;

            return result;
        }
    }
}
