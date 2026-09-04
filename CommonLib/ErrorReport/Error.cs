using System;
using System.Collections.Generic;
using System.Diagnostics;

using CommonLib.ErrorReport.Base;

namespace CommonLib.ErrorReport
{
    [DebuggerDisplay("{SheetName} : {Message}")]
    [Serializable]
    public class Error
    {
        private string _sheetName;
        public string SheetName
        {
            get { return string.IsNullOrEmpty(_sheetName) ? "" : _sheetName.Length > 31 ? _sheetName.Substring(0, 31) : _sheetName; }
            set { _sheetName = value; }
        }
        public object ErrorType { set; get; }
        public string Link
        {
            get
            {
                return GetHyperlinkFomula(SheetName, RowNum, ColNum, "Link");
            }
        }

        public ErrorLevel ErrorLevel { get; set; }
        public int RowNum { get; set; }
        public int ColNum { get; set; }
        public string Message { get; set; }

        public List<string> Comments = new List<string>();

        private string GetHyperlinkFomula(string sheetName, int row, int column, string friendlyName)
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
                return ("$" + GetColumnLetter(column) + "$" + row);
            }

            return (GetColumnLetter(column) + row);
        }


        public string GetAddress()
        {
            if (RowNum == 0 || ColNum == 0)
            {
                return "#REF!";
            }

            return (GetColumnLetter(ColNum) + RowNum);
        }

        private static string GetColumnLetter(int iColumnNumber)
        {
            if (iColumnNumber < 1)
            {
                return "#REF!";
            }

            string sCol = "";
            do
            {
                sCol = ((char)('A' + ((iColumnNumber - 1) % 26))) + sCol;
                iColumnNumber = (iColumnNumber - ((iColumnNumber - 1) % 26)) / 26;
            }
            while (iColumnNumber > 0);
            return sCol;
        }
    }
}
