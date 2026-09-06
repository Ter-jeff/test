using System.Collections.Generic;
using System.Diagnostics;

namespace CommonLib.ErrorReport.Base
{
    [DebuggerDisplay("{SheetName} : {Message}")]
    public class Error
    {
        private string _sheetName = "";
        public string SheetName
        {
            get { return string.IsNullOrEmpty(_sheetName) ? "" : _sheetName; }
            set { _sheetName = value; }
        }
        public string Link
        {
            get
            {
                return GetHyperlinkFormula(SheetName, RowNum, ColNum, "Link");
            }
        }

        public ErrorCode ErrorCode { get; } = null!;
        public EnumErrorLevel ErrorLevel { get; }
        public int RowNum { get; }
        public int ColNum { get; }
        public string[]? MessageArgs { get; }

        public string ColLetter
        {
            get
            {
                string letter = GetColumnLetter(ColNum);
                if (letter == "#REF!")
                {
                    return "";
                }

                return letter;
            }
        }

        public string Message { get; } = string.Empty;
        public ErrorInfo? ErrorInfo { get; private set; }
        public string? Pattern { get; set; }

        public List<string> Comments = [];

        public Error(Error error)
        {
            if (error == null)
            {
                return;
            }

            _sheetName = error._sheetName;
            ErrorCode = error.ErrorCode;
            ErrorLevel = error.ErrorLevel;
            RowNum = error.RowNum;
            ColNum = error.ColNum;
            Message = error.Message;
            Pattern = error.Pattern;

            if (error.Comments != null)
            {
                Comments = [.. error.Comments];
            }
        }

        public Error(ErrorCode errorCode, EnumErrorLevel enumErrorLevel, string sheetName, int rowNum, int colNum, string message, ErrorInfo? errorInfo = null)
        {
            ErrorCode = errorCode;
            ErrorLevel = enumErrorLevel;
            SheetName = sheetName;
            RowNum = rowNum;
            ColNum = colNum;
            Message = message;
            if (errorInfo != null)
            {
                ErrorInfo = errorInfo;
                Pattern = errorInfo.Pattern;
                Comments = errorInfo.Comments;
            }
        }

        public Error(ErrorCode errorCode, string sheetName, int rowNum, int colNum, string[]? messageArgs = null, ErrorInfo? errorInfo = null)
        {
            ErrorCode = errorCode;
            ErrorLevel = errorCode.ErrorLevel;
            SheetName = sheetName;
            RowNum = rowNum;
            ColNum = colNum;
            Message = errorCode.FormatMessage(messageArgs);
            MessageArgs = messageArgs;
            if (errorInfo != null)
            {
                ErrorInfo = errorInfo;
                Pattern = errorInfo.Pattern;
                Comments = errorInfo.Comments;
            }
        }

        public Error Copy()
        {
            return new Error(this);
        }

        private static string GetHyperlinkFormula(string sheetName, int row, int column, string friendlyName)
        {
            //if (row == 0)
            //    return "";
            if (column == 0)
            {
                return "=HYPERLINK(\"#\'" + sheetName + "\'!" + row + ":" + row + "\",\"" + friendlyName + "\")";
            }

            return "=HYPERLINK(\"#\'" + sheetName + "\'!" + GetAddress(row, column) + "\",\"" + friendlyName + "\")";
        }

        public string GetHyperlink()
        {
            if (ColNum == 0)
            {
                return SheetName + "!" + RowNum + ":" + RowNum;
            }

            return SheetName + "!" + GetAddress(RowNum, ColNum);
        }

        public static string GetAddress(int row, int column, bool absolute = false)
        {
            if (row == 0 || column == 0)
            {
                return "#REF!";
            }

            if (absolute)
            {
                return "$" + GetColumnLetter(column) + "$" + row;
            }

            return GetColumnLetter(column) + row;
        }

        public string GetAddress()
        {
            if (RowNum == 0 || ColNum == 0)
            {
                return "#REF!";
            }

            return GetColumnLetter(ColNum) + RowNum;
        }

        private static string GetColumnLetter(int columnNumber)
        {
            if (columnNumber < 1)
            {
                return "#REF!";
            }

            string columnLetter = "";
            do
            {
                columnLetter = (char)('A' + ((columnNumber - 1) % 26)) + columnLetter;
                columnNumber = (columnNumber - ((columnNumber - 1) % 26)) / 26;
            }
            while (columnNumber > 0);
            return columnLetter;
        }

        public string Print()
        {
            return $"Sheet: {SheetName}, Row: {RowNum}, Col: {ColLetter}, " +
                   $"ErrorCode: {ErrorCode.FullCode}, Level: {ErrorLevel}, " +
                   $"Category: {ErrorCode.EnumErrorCategory}, " +
                   $"Behavior: {ErrorCode.EnumErrorBehavior?.ToString()}, " +
                   $"Target: {ErrorCode.EnumErrorTarget?.ToString()}, " +
                   $"Message: {Message}, Guidance: {ErrorCode.Guidance}, Link: {Link}, ";
        }
    }
}
