using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class IllegalCharChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            var regexWord = new Regex("^[0-9a-zA-Z]+$", RegexOptions.Singleline | RegexOptions.Compiled);
            var regexUserDef8 = new Regex("^[0-9a-zA-Z]+_?$", RegexOptions.Singleline | RegexOptions.Compiled);  // user_define 8 allow "X_"

            foreach (Characterization item in charList)
            {
                // only check use item
                if (!Regex.IsMatch(item.Use, "^use$", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                string errorMessage = "";
                var illegalList = new List<string>();

                errorMessage += _check(regexWord, item.UserDef1, "UserDef1 ", illegalList);
                errorMessage += _check(regexWord, item.UserDef2, "UserDef2 ", illegalList);
                errorMessage += _check(regexWord, item.UserDef3, "UserDef3 ", illegalList);
                errorMessage += _check(regexWord, item.UserDef4, "UserDef4 ", illegalList);
                errorMessage += _check(regexWord, item.UserDef5, "UserDef5 ", illegalList);
                errorMessage += _check(regexWord, item.UserDef6, "UserDef6 ", illegalList);
                errorMessage += _check(regexWord, item.UserDef7, "UserDef7 ", illegalList);
                errorMessage += _check(regexUserDef8, item.UserDef8, "UserDef8 ", illegalList);
                errorMessage += _check(regexWord, item.Category, "Category ", illegalList);
                errorMessage += _check(regexWord, item.Group, "Group ", illegalList);

                if (errorMessage == "")
                {
                    continue;
                }

                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Error,
                    ErrorType = ErrorType.IllegalChar,
                    SheetName = item.SheetName,
                    RowNum = item.RowNum,
                    Message = "There is illegal char in " + errorMessage,
                    CommentsList = illegalList.Distinct().ToList(),
                });
                ErrorReportManager.AddError(CharErrorType.E_IllegalChar_01, item.SheetName, item.RowNum, 0, [errorMessage], new ErrorInfo() { Comments = illegalList.Distinct().ToList() });
            }
        }

        private static string _check(Regex reg, string valStr, string field, ICollection<string> illegalList)
        {
            if (string.IsNullOrEmpty(valStr) || reg.Match(valStr).Success)
            {
                return "";
            }

            illegalList.Add(valStr);
            return field;
        }
    }
}
