using System.Collections.Generic;
using System.Linq;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.BinCut.Flow;

namespace TestPlanLib.BinCut.Checker
{
    public static class ValidateBinCutPost
    {
        public static void CheckFlowSheetsPmode(BinCutFlowSheets binCutFlowSheets)
        {
            for (int i = 0; i < binCutFlowSheets.Count; i++)
            {
                var sheet1Pmodes = binCutFlowSheets[i].First().Rows.Where(x => !string.IsNullOrEmpty(x.PerformanceMode)).Select(y => y.PerformanceMode).ToList();
                for (int j = i + 1; j < binCutFlowSheets.Count; j++)
                {
                    List<BinCutFlowSheetRow> sheet2Rows = binCutFlowSheets[j].First().Rows;
                    foreach (BinCutFlowSheetRow row in sheet2Rows)
                    {
                        if (sheet1Pmodes.Any(x => x.EqualsIgnoreCase(row.PerformanceMode)))
                        {
                            string errorMessage = $"The performance {row.PerformanceMode} is duplicated in other post flow sheets !!!";
                            binCutFlowSheets[j].First().AddError(BinCutErrorType.E_PerformanceMode_01, binCutFlowSheets[j].First().SheetName, row.RowNum, binCutFlowSheets[j].First().Indices.PerformanceModeIndex, $"The performance {row.PerformanceMode} is duplicated in other post flow sheets !!!", [row.PerformanceMode]);
                        }
                    }
                }
            }
        }
    }
}
