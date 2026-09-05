using System.Collections.Generic;
using System.Linq;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace CommonReaderLib
{
    public abstract class MySheet
    {
        private readonly List<Error> _errors = [];
        public string SheetName = string.Empty;

        protected MySheet()
        {
        }

        protected MySheet(MySheet mySheet)
        {
            if (mySheet == null)
            {
                return;
            }

            SheetName = mySheet.SheetName;
            _errors = mySheet._errors?.Select(row => row.Copy()).ToList() ?? [];
        }

        public List<Error> GetErrors()
        {
            return _errors;
        }

        public void ClearErrors()
        {
            _errors.Clear();
        }

        public void AddError(ErrorCode errorCode, string sheetName, int rowNum, int colNum, string message, ErrorInfo? errorInfo = null)
        {
            var error = new Error(errorCode, errorCode.ErrorLevel, sheetName, rowNum, colNum, message, errorInfo);
            _errors.Add(error);
        }

        public void AddError(ErrorCode errorCode, string sheetName, int rowNum, int colNum, string message, string[] messageArgs, ErrorInfo? errorInfo = null)
        {
            var error = new Error(errorCode, sheetName, rowNum, colNum, messageArgs, errorInfo);
            _errors.Add(error);
        }

        public void AddError(ErrorCode errorCode, string sheetName, int rowNum, int colNum, string[]? messageArgs = null, ErrorInfo? errorInfo = null)
        {
            var error = new Error(errorCode, sheetName, rowNum, colNum, messageArgs, errorInfo);
            _errors.Add(error);
        }

        public void AddToErrorReport()
        {
            foreach (Error error in _errors)
            {
                ErrorReportManager.AddError(error.ErrorCode, error.SheetName, error.RowNum, error.ColNum, error.Message, error.ErrorInfo);
            }
        }

        public void AddDimensionError()
        {
            AddError(PreActionErrorType.E_InvalidDocument_01, SheetName, 0, 0);
        }

        public void AddFirstHeaderError(string firstHeader)
        {
            AddError(PreActionErrorType.E_MissingHeader_01, SheetName, 0, 0, [firstHeader]);
        }
    }
}
