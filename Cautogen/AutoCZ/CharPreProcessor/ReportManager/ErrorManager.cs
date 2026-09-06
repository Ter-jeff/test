using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;

namespace Cautogen.AutoCZ.CharPreProcessor.ReportManager
{
    public class ErrorManager
    {
        public static Dictionary<ErrorType, List<ErrorMessage>> ErrorListDict;

        /* Generic Add */
        public static void Add(ErrorMessage errMessage)
        {
            ErrorType errorType = errMessage.ErrorType;

            if (ErrorListDict.ContainsKey(errorType))
            {
                ErrorListDict[errorType].Add(errMessage);
            }
            else
            {
                ErrorListDict[errorType] = new List<ErrorMessage> { errMessage };
            }
        }

        private static void _Add(ErrorType errorType, ErrorLevel errorLevel, string sheetName, int rowNum, List<int> colNums, string use, string message, params string[] commentStrings)
        {
            // only record the used one
            if (string.IsNullOrEmpty(use) || !Regex.IsMatch(use, "^use$", RegexOptions.IgnoreCase))
            {
                return;
            }

            Add(new ErrorMessage
            {
                SheetName = sheetName,
                RowNum = rowNum,
                ColList = colNums,
                Message = message,
                ErrorLevel = errorLevel,
                CommentsList = commentStrings.ToList(),
                ErrorType = errorType,
            });
        }

        /* Add Error */
        public static void AddError(ErrorType errorType, string sheetName, int rowNum, List<int> colNums, string use, string message, params string[] commentStrings)
        {
            _Add(errorType, ErrorLevel.Error, sheetName, rowNum, colNums, use, message, commentStrings);
        }
        public static void AddError(ErrorType errorType, string sheetName, int rowNum, int colNum, string use, string message, params string[] commentStrings)
        {
            _Add(errorType, ErrorLevel.Error, sheetName, rowNum, new List<int> { colNum }, use, message, commentStrings);
        }
        //todo
        public static void AddError(ErrorType errorType, Characterization item, List<int> col, string message, params string[] commentStrings)
        {
            _Add(errorType, ErrorLevel.Error, item.SheetName, item.RowNum, col, item.Use, message, commentStrings);
        }


        /* Add Warning */
        public static void AddWarning(ErrorType errorType, string sheetName, int rowNum, List<int> colNums, string use, string message,
            params string[] commentStrings)
        {
            _Add(errorType, ErrorLevel.Warning, sheetName, rowNum, colNums, use, message, commentStrings);
        }
        //todo
        public static void AddWarning(ErrorType errorType, Characterization item, List<int> colNums, string message,
            params string[] comments)
        {
            _Add(errorType, ErrorLevel.Warning, item.SheetName, item.RowNum, colNums, item.Use, message, comments);
        }
    }
}
