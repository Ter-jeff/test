using System.Collections.Generic;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class Userder1Checker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization item in charList)
            {
                if (CharPlan.HardIpSheets.Contains(item.SheetName.ToUpper()))
                {
                    if (Regex.IsMatch(item.UserDef1, "HAC|HFL|HFH|HFLH|HIO", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    ErrorMessages.Add(new ErrorMessage()
                    {
                        ErrorLevel = ErrorLevel.Error,
                        ErrorType = ErrorType.WrongUserdef1,
                        SheetName = item.SheetName,
                        RowNum = item.RowNum,
                        Message = "Wrong USERDEF1 in HardIP sheets in " + item.SheetName + "  Row: " + item.RowNum,
                    });
                    ErrorReportManager.AddError(CharErrorType.E_WrongUserdef1_01, item.SheetName, item.RowNum, 0, ["HardIP", item.SheetName, $"{item.RowNum}"]);
                }
                else
                {
                    if (Regex.IsMatch(item.UserDef1, "DFTL|DFTH|DFTLH|MCL|MCH|MCLH"))
                    {
                        continue;
                    }

                    ErrorMessages.Add(new ErrorMessage()
                    {
                        ErrorLevel = ErrorLevel.Error,
                        ErrorType = ErrorType.WrongUserdef1,
                        SheetName = item.SheetName,
                        RowNum = item.RowNum,
                        Message = "Wrong USERDEF1 in non-HardIP sheets in " + item.SheetName + "  Row: " + item.RowNum,
                    });
                    ErrorReportManager.AddError(CharErrorType.E_WrongUserdef1_01, item.SheetName, item.RowNum, 0, ["non-HardIP", item.SheetName, $"{item.RowNum}"]);
                }
            }
        }
    }
}
