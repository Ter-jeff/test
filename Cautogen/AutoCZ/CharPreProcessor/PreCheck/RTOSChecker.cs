using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class RtosChecker : PreCheckBase
    {
        private List<string> _keyList;

        public override void Check(List<Characterization> charList, string sheetName)
        {
            // check only if IsUseRtosCmd is enabled
            if (UtilityMain.UtilityData == null)
            {
                return;
            }

            if (!UtilityMain.UtilityData.InputParam.IsUseRtosCmd)
            {
                return;
            }

            if (string.IsNullOrEmpty(CharPlan.Revision.RtosCmdPattern))
            {
                return;
            }

            _keyList = GetKeyList(CharPlan.Revision.RtosCmdPattern);

            // gating only on both used and useRtosCmd rows
            foreach (Characterization item in charList.Where(i => IsUseItem(i) && i.IsUseRtosCmd)
                                         .Where(i => !_CheckRtosSyntaxPass(i.UserDef6)))
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Error,
                    ErrorType = ErrorType.RtosUserdef6Syntax,
                    SheetName = sheetName,
                    RowNum = item.RowNum,
                    Message = $"RTOS_Userdef6_Syntax: {item.UserDef6} is wrong",
                    CommentsList = new List<string> { },
                });
                ErrorReportManager.AddError(CharErrorType.E_RtosUserdef6Syntax_02, sheetName, item.RowNum, 0, [item.UserDef6],
                    new ErrorInfo() { Comments = new List<string> { } });
            }
        }

        private static List<string> GetKeyList(string rtosCmdPattern)
        {
            string[] list = rtosCmdPattern.Split(' ');
            return list.Where(item => item.StartsWith("$", StringComparison.OrdinalIgnoreCase)).Select(y => y.Substring(1)).ToList();
        }

        private bool _CheckRtosSyntaxPass(string userdef6)
        {
            if (userdef6.Equals("X", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string[] tokens = userdef6.ToLower().Split(new[] { "nbsp" }, StringSplitOptions.RemoveEmptyEntries);

            bool isPass = tokens.All(token => _keyList.Any(key => token.StartsWith(key, StringComparison.OrdinalIgnoreCase)));
            return isPass;
        }
    }
}
