using System.Collections.Generic;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class PatternCellChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization charRow in charList)
            {
                if (UtilityMain.UtilityData.InputParam.CharPreCheckForNewTChar)
                {
                    CheckForNewTChar(charRow, sheetName);
                }
            }
        }
        private void CheckForNewTChar(Characterization charRow, string sheetName)
        {
            CheckMergedPattern(charRow, sheetName);
        }
        private void CheckMergedPattern(Characterization charRow, string sheetName)
        {
            foreach (PatternCell patternCell in charRow.PatternCellList)
            {
                string[] splitLength = patternCell.PatternDefine.Replace(" ", "").Split(',');
                if (splitLength.Length > 1)
                {
                    ErrorMessages.Add(new ErrorMessage
                    {
                        ErrorLevel = ErrorLevel.Error,
                        ErrorType = ErrorType.IllegalForNewTChar,
                        SheetName = sheetName,
                        RowNum = charRow.RowNum,
                        ColList = new List<int> { patternCell.ColIndex },
                        Message = "Can't define multi patterns here, please define one cell by one pattern.",
                        CommentsList = new List<string> { },
                    });
                    ErrorReportManager.AddError(CharErrorType.E_IllegalForNewTChar_01, sheetName, charRow.RowNum, patternCell.ColIndex, [],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
        }
    }
}
