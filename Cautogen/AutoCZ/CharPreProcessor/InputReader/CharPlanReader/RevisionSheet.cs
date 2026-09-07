using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader
{
    public class RevisionSheet
    {
        /* properties */
        public string DefaultSrcValue;
        public string RtosCmdPattern;
        private string _sheetName;
        private bool _hasRevisionSheet;
        private bool _ismissingDigSrc;

        /* constructor */
        public RevisionSheet()
        {
            DefaultSrcValue = "0";
            _sheetName = "Revision";
            _hasRevisionSheet = false;
            _ismissingDigSrc = false;
        }

        public void Read(ExcelWorksheet sh)
        {
            _sheetName = sh.Name;
            _hasRevisionSheet = true;
            _ismissingDigSrc = true;  // assume not found
            _Search_Settings(sh);
        }

        private void _Search_Settings(ExcelWorksheet sh)
        {
            int colDim = sh.Dimension.End.Column >= 50 ? 50 : sh.Dimension.End.Column;
            for (int i = 1; i <= sh.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= colDim; j++)
                {
                    object value = sh.Cells[i, j].Value;

                    if (value == null)
                    {
                        continue;
                    }

                    // search for "sgmt_default = 1" or "sgmt_default = 0" in each cell
                    if (Regex.IsMatch(value.ToString(), @"sgmt_default\s*=\s*\d", RegexOptions.IgnoreCase))
                    {
                        DefaultSrcValue = Regex.Match(value.ToString(), @"sgmt_default\s*=\s*(?<value>\d)", RegexOptions.IgnoreCase).Groups["value"].ToString();
                        _ismissingDigSrc = false;
                    }

                    // search for RTOS_Userdef6_Syntax
                    if (Regex.IsMatch(value.ToString(), "RTOS_Userdef6_Syntax:", RegexOptions.IgnoreCase))
                    {
                        RtosCmdPattern = Regex.Match(value.ToString(), "RTOS_Userdef6_Syntax:(?<syntax>.*)").Groups["syntax"].ToString().Trim();
                    }
                }
            }
        }

        public void Check()
        {
            // report warning if no revision sheet in char plan
            if (!_hasRevisionSheet)
            {
                ErrorManager.AddWarning(ErrorType.MissingRevisonSheet, "", 0, new List<int>(0), "Use", "Missing revision sheet in CharPlan -- need use Revision sheet to specify default digsrc value");
                foreach (int col in new List<int>(0))
                {
                    ErrorReportManager.AddError(CharErrorType.W_MissingRevisonSheet_01, "", 0, col, []);
                }
            }

            // report warning if DigScr default value is not assinged
            if (_ismissingDigSrc)
            {
                const string message = "sgmt default value is not found, use '0' as sgmt default value," + "\r\n" +
                    "If wish to define the sgmt default value, just add 'sgmt_default=0'" + "\r\n" +
                    "or 'sgmt_default=1' in the Revision sheet.";
                ErrorManager.AddWarning(ErrorType.WrongDigSrc, _sheetName, 1, new List<int>(0), "Use", message);
                foreach (int col in new List<int>(0))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongDigSrc_16, _sheetName, 1, col, []);
                }
            }

            // report warning if RTOS Userdef6 Syntax is not assigned
            if (UtilityMain.UtilityData.InputParam.IsUseRtosCmd && string.IsNullOrEmpty(RtosCmdPattern))
            {
                ErrorManager.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Warning,
                    ErrorType = ErrorType.RtosUserdef6Syntax,
                    SheetName = "Revision",
                    RowNum = 3,
                    Message = "There is no RTOS_Userdef6_Syntax in Revision sheet",
                    CommentsList = new List<string> { },
                });
                ErrorReportManager.AddError(CharErrorType.W_RtosUserdef6Syntax_01, "Revision", 3, 0, [],
                    new ErrorInfo() { Comments = new List<string> { } });
            }
        }
    }
}
