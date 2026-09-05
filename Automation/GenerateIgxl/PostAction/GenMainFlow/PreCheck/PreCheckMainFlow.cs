using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow.PreCheck
{
    public class PreCheckMainFlow
    {
        #region protected variable
        private readonly ExcelWorkbook _baseWorkbook;
        private ExcelWorksheet _checkSheet;
        private string _sheetName;
        private const int MaxStartColumn = 10;
        private const int MaxStartRow = 10;
        #endregion

        public PreCheckMainFlow(ExcelWorkbook baseWorkbook)
        {
            _baseWorkbook = baseWorkbook;
        }

        #region public method
        public void SetSheetName(string sheetName)
        {
            _sheetName = sheetName;
        }

        public bool CheckMain()
        {
            bool checkResult = CheckExist();

            if (!checkResult)
            {
                return false;
            }

            checkResult = CheckFormat();

            if (!checkResult)
            {
                return false;
            }

            checkResult = CheckBusiness();

            if (!checkResult)
            {
                return false;
            }

            return true;
        }

        #endregion

        #region private method
        private bool CheckExist()
        {
            bool checkResult = CheckSheetName();
            if (!checkResult)
            {
                return false;
            }

            checkResult = CheckFirstDataLocation();
            if (!checkResult)
            {
                return false;
            }

            checkResult = CheckHeader();
            if (!checkResult)
            {
                return false;
            }

            return true;
        }

        private bool CheckSheetName()
        {
            bool checkResult = _baseWorkbook.IsExist(_sheetName);
            if (checkResult)
            {
                if (_baseWorkbook.IsExist(_sheetName))
                {
                    _checkSheet = _baseWorkbook.Worksheets[_sheetName];
                }
            }
            else
            {
                ErrorReportManager.AddError(FlowMainErrorType.E_MissingDocument_01, _sheetName, 1, 1, [_sheetName]);
                return false;
            }

            return true;
        }

        private bool CheckFirstDataLocation()
        {
            for (int j = 1; j <= MaxStartColumn; j++)
            {
                for (int i = 1; i <= MaxStartRow; i++)
                {
                    string content = _checkSheet.GetCellValue(i, j);
                    if (!string.IsNullOrEmpty(content))
                    {
                        return true;
                    }
                }
            }

            ErrorReportManager.AddError(FlowMainErrorType.E_InvalidFormat_01, _sheetName, 1, 1);
            return false;
        }

        #endregion

        #region Private Check Method

        private bool CheckHeader()
        {
            return true;
        }

        private bool CheckFormat()
        {
            return true;
        }

        private bool CheckBusiness()
        {

            return true;
        }
        #endregion
    }
}
