using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    internal class AtLeastOnePayloadChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charRows, string sheetName)
        {
            // For non-Rtos cmd char row, there should be at least one payload in the patterns
            foreach (Characterization charRow in charRows.Where(row => !row.HasAnyPayload && !row.IsUseRtosCmd))
            {
                string outString = "There is no payload in these patterns: " + string.Join(",", charRow.NonEmptyPatterns.Values);
                ErrorManager.AddError(ErrorType.PatternsWithoutPayload, charRow.SheetName,
                    charRow.RowNum, charRow.ColNum(charRow.NonEmptyPatterns.Keys.ToList()), "Use", outString);
                foreach (int col in charRow.ColNum(charRow.NonEmptyPatterns.Keys.ToList()))
                {
                    ErrorReportManager.AddError(CharErrorType.E_PatternsWithoutPayload_01, charRow.SheetName, charRow.RowNum, col,
                        [string.Join(",", charRow.NonEmptyPatterns.Values)]);
                }
            }
        }
    }
}
