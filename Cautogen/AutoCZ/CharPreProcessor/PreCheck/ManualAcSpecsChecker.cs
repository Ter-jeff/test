using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

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
    public class ManualAcSpecsChecker : PreCheckBase
    {
        private Dictionary<string, Dictionary<string, string>> _manualAcSheet;
        public override void Check(List<Characterization> charList, string sheetName)
        {
            if (CharPlan.OptionalTimesetting == null)
            {
                return;
            }
            else
            {
                _manualAcSheet = CharPlan.OptionalTimesetting.AcCateSymbolValue;
            }
            _CheckManualAcInProgram();

            foreach (Characterization charRow in charList)
            {
                if (charRow.TimeSet.Split(':').Length < 2)
                {
                    continue;
                }

                _CheckManualAcInTimeSettingSheet(charRow);
            }
        }

        private void _CheckManualAcInProgram()
        {
            int col = 1;
            foreach (KeyValuePair<string, Dictionary<string, string>> manualAc in _manualAcSheet)
            {
                if (!UtilityMain.UtilityData.AcCategories.Any(x => string.Equals(manualAc.Key, x, StringComparison.OrdinalIgnoreCase)))
                {
                    string searchBlock = SearchBlockAc(manualAc.Key);
                    searchBlock = searchBlock == "" ? "None" : searchBlock;
                    ErrorMessages.Add(new ErrorMessage
                    {
                        ErrorLevel = ErrorLevel.Warning,
                        ErrorType = ErrorType.MissingManualAcInProgram,
                        SheetName = "timesettings",
                        RowNum = 2,
                        ColList = new List<int>() { col },
                        Message =
                            $"Can't serach category: {manualAc.Key} in program ac specs, use {searchBlock} as base to generate.",
                        CommentsList = new List<string> { },
                    });
                    ErrorReportManager.AddError(CharErrorType.W_MissingManualAcInProgram_01, "timesettings", 2, col, [manualAc.Key, searchBlock],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
                col++;
            }
        }

        private void _CheckManualAcInTimeSettingSheet(Characterization charRow)
        {
            string manualAcPlan = charRow.TimeSet.Split(':')[1];
            if (!string.IsNullOrEmpty(charRow.ShiftFreq))
            {
                manualAcPlan += "_" + charRow.ShiftFreq;
                if (!manualAcPlan.EndsWith("MHz", StringComparison.OrdinalIgnoreCase))
                {
                    manualAcPlan += "MHz";
                }
            }

            if (!_manualAcSheet.ContainsKey(manualAcPlan) &&
                !UtilityMain.UtilityData.AcCategories.Any(x => Regex.IsMatch(manualAcPlan, x, RegexOptions.IgnoreCase)))
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Error,
                    ErrorType = ErrorType.MissingManualAcInProgAndTsSheet,
                    SheetName = charRow.SheetName,
                    RowNum = charRow.RowNum,
                    ColList = charRow.ColNum("timeset"),
                    Message = $"Char manual ac: {manualAcPlan} is not in timesettings sheet and program ac spec.",
                    CommentsList = new List<string> { },
                });
                foreach (int col in charRow.ColNum("timeset"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_MissingManualAcInProgAndTsSheet_01,
                        charRow.SheetName, charRow.RowNum, col, [manualAcPlan],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
            else if (!_manualAcSheet.ContainsKey(manualAcPlan))
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Warning,
                    ErrorType = ErrorType.MissingManualAcInTimeSettingsSheet,
                    SheetName = charRow.SheetName,
                    RowNum = charRow.RowNum,
                    ColList = charRow.ColNum("timeset"),
                    Message =
                        $"Char manual ac: {manualAcPlan} is not in timesettings sheet, directly use program ac spec.",
                    CommentsList = new List<string> { },
                });
                foreach (int col in charRow.ColNum("timeset"))
                {
                    ErrorReportManager.AddError(CharErrorType.W_MissingManualAcInTimeSettingsSheet_01,
                        charRow.SheetName, charRow.RowNum, col, [manualAcPlan],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
        }

        private string SearchBlockAc(string manualCate)
        {
            string result = "";
            string manualFreq = Regex.Match(manualCate, @"\d+Mhz").Value;
            foreach (string cate in UtilityMain.UtilityData.AcCategories)
            {
                if (!string.IsNullOrEmpty(Regex.Match(manualCate, cate, RegexOptions.IgnoreCase).Value))
                {
                    result = Regex.Match(manualCate, cate, RegexOptions.IgnoreCase).Value;
                    break;
                }
            }

            if (!string.IsNullOrEmpty(result) && !string.IsNullOrEmpty(manualFreq) &&
                UtilityMain.UtilityData.AcCategories.Any(x => string.Equals(x, result + "_" + manualFreq, StringComparison.OrdinalIgnoreCase)))
            {
                result =
                    UtilityMain.UtilityData.AcCategories.FirstOrDefault(
                        x => string.Equals(x, result + "_" + manualFreq, StringComparison.OrdinalIgnoreCase));
            }

            return result;
        }
    }
}
