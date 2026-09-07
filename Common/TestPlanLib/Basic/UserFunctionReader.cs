using System;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Basic
{
    public class UserFunctionReader : MySheetReader<UserFunctionSheet>
    {
        private const string ConHeaderUserFunction = "UserFunction";
        private const string ConHeaderArgumentSetting = "ArgumentSetting";
        private const string ConHeaderMultiFstpSetting = "MultiFstpSetting";
        private const string ConHeaderCallAfterInstance = "CallAfterInstance";
        private int _indexUserFunction = -1;
        private int _indexArgumentSetting = -1;
        private int _indexMultiFstpSetting = -1;
        private int _indexCallAfterInstance = -1;
        private readonly UserFunctionSheet _userFunctionSheet = new();

        public override UserFunctionSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            ReadSheet();

            return _userFunctionSheet;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow;
            int colNum = EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderUserFunction))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderUserFunction))
                {
                    _indexUserFunction = i;
                    continue;
                }

                if (header.EqualsIgnoreCase(ConHeaderArgumentSetting))
                {
                    _indexArgumentSetting = i;
                    continue;
                }

                if (header.EqualsIgnoreCase(ConHeaderMultiFstpSetting))
                {
                    _indexMultiFstpSetting = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderCallAfterInstance))
                {
                    _indexCallAfterInstance = i;
                    continue;
                }
            }
        }

        private void ReadSheet()
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var rowTmp = new UserFunctionTableRow();
                if (_indexUserFunction != -1)
                {
                    rowTmp.UserFunction = ExcelWorksheet.GetCellValue(i, _indexUserFunction).Trim();
                }
                if (_indexCallAfterInstance != -1)
                {
                    rowTmp.CallAfterInstance = ExcelWorksheet.GetCellValue(i, _indexCallAfterInstance).Trim();
                }
                if (_indexArgumentSetting != -1)
                {
                    string argumentListString = ExcelWorksheet.GetCellValue(i, _indexArgumentSetting).Trim();
                    string[] argumentList = argumentListString.Split([";"], StringSplitOptions.RemoveEmptyEntries);
                    foreach (string argument in argumentList)
                    {
                        int sep = argument.IndexOf('=');
                        if (sep < 0)
                        {
                            continue;
                        }

                        string argumentKey = argument[..sep].Trim();
                        string argumentValue = argument[(sep + 1)..].Trim();
                        rowTmp.ArgumentSetting.Add(argumentKey, argumentValue);
                    }
                }
                if (_indexMultiFstpSetting != -1)
                {
                    rowTmp.MultiFstpSetting = ExcelWorksheet.GetCellValue(i, _indexMultiFstpSetting).Trim();
                }
                _userFunctionSheet.Rows.Add(rowTmp);
            }
        }
    }
}
