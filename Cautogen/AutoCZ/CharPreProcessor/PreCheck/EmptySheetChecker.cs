using System.Collections.Generic;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class EmptySheetChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            if (charList.Count == 0)
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Warning,
                    ErrorType = ErrorType.EmptySheet,
                    SheetName = sheetName,
                    RowNum = 1,
                    Message = $"sheet {sheetName} is empty, all items are unused! "
                });
                ErrorReportManager.AddError(CharErrorType.W_EmptySheet_01, sheetName, 1, 0, [sheetName]);
            }
        }
    }
}
