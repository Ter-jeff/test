using System.Collections.Generic;
using System.Linq;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class DuplicatePatternChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            /* check each pattern only show once in a singel char row*/
            foreach (Characterization item in charList.Where(IsUseItem))
            {
                var patterns = item.AllPatterns.Values.ToList();

                var duplicatedPatterns =
                    patterns.GroupBy(x => x).Where(p => !string.IsNullOrEmpty(p.Key) && p.Count() > 1).Select(p => p.Key).ToList();

                if (duplicatedPatterns.Count == 0)
                {
                    continue;
                }

                var duplicatedPatHeaders = item.AllPatterns.Where(p => duplicatedPatterns.Contains(p.Value)).Select(p => p.Key).ToList();
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Warning,
                    ErrorType = ErrorType.DuplicatePattern,
                    SheetName = item.SheetName,
                    RowNum = item.RowNum,
                    ColList = item.ColNum(duplicatedPatHeaders),
                    Message = "duplicate patterns",
                    CommentsList = new List<string> { string.Join(",", duplicatedPatterns) },
                });

                foreach (int col in item.ColNum(duplicatedPatHeaders))
                {
                    ErrorReportManager.AddError(CharErrorType.E_DuplicatePattern_01, item.SheetName, item.RowNum, col, [],
                        new ErrorInfo() { Comments = new List<string> { string.Join(",", duplicatedPatterns) } });
                }
            }
        }
    }
}
