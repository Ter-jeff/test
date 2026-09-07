using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class ProgramMappingResultChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization charRow in charList)
            {
                if (charRow.MappingKey == null || charRow.MappingKey.Payload == null)
                {
                    continue;
                }

                List<int> mappingPatternIndex = charRow.ColNum(charRow.AllPatterns.FirstOrDefault(x => x.Value.IndexOf(charRow.MappingKey.Payload, StringComparison.OrdinalIgnoreCase) != -1).Key);
                _CheckMappingAcCategory(charRow, sheetName, mappingPatternIndex);
                _CheckMappingTimeset(charRow, sheetName, mappingPatternIndex);
                _CheckMappingDcCategory(charRow, sheetName, mappingPatternIndex);
                //_CheckMappingLevel(charRow, sheetName, mappingPatternIndex);
            }
        }

        private void _CheckMappingAcCategory(Characterization charRow, string sheetName, List<int> col)
        {
            if (charRow.TimeSet.Split(':').Length > 1)
            {
                if (!string.IsNullOrEmpty(charRow.TimeSet.Split(':')[1].Trim()))
                {
                    return;
                }
            }

            if (charRow.MappingSpec.AcCategory.Count > 1)
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Warning,
                    ErrorType = ErrorType.MultiUsedAcForPayload,
                    SheetName = sheetName,
                    RowNum = charRow.RowNum,
                    ColList = col,
                    Message =
                        $"Multi ac category for payload({charRow.MappingKey.Payload}) in base program: {string.Join(",", charRow.MappingSpec.AcCategory)} .",
                    CommentsList = new List<string> { },
                });
                foreach (int column in col)
                {
                    ErrorReportManager.AddError(CharErrorType.W_MultiUsedAcForPayload_01, sheetName, charRow.RowNum, column,
                        [charRow.MappingKey.Payload, string.Join(",", charRow.MappingSpec.AcCategory)],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
            if (charRow.MappingSpec.AcCategory.Count < 1)
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Error,
                    ErrorType = ErrorType.MissingUsedAcForPayload,
                    SheetName = sheetName,
                    RowNum = charRow.RowNum,
                    ColList = col,
                    Message = $"None of ac category for payload({charRow.MappingKey.Payload}) in base program.",
                    CommentsList = new List<string> { },
                });
                foreach (int column in col)
                {
                    ErrorReportManager.AddError(CharErrorType.E_MissingUsedAcForPayload_01, sheetName, charRow.RowNum, column,
                        [charRow.MappingKey.Payload],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
        }
        private void _CheckMappingTimeset(Characterization charRow, string sheetName, List<int> col)
        {
            if (!string.IsNullOrEmpty(charRow.TimeSet.Split(':')[0].Trim()))
            {
                return;
            }

            if (charRow.MappingSpec.Timeset.Count > 1)
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Warning,
                    ErrorType = ErrorType.MultiUsedTimesetForPayload,
                    SheetName = sheetName,
                    RowNum = charRow.RowNum,
                    ColList = col,
                    Message =
                        $"Multi timeset for payload({charRow.MappingKey.Payload}) in base program: {string.Join(",", charRow.MappingSpec.Timeset)} .",
                    CommentsList = new List<string> { },
                });
                foreach (int column in col)
                {
                    ErrorReportManager.AddError(CharErrorType.W_MultiUsedTimesetForPayload_01, sheetName, charRow.RowNum, column,
                        [charRow.MappingKey.Payload, string.Join(",", charRow.MappingSpec.Timeset)],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
            if (charRow.MappingSpec.Timeset.Count < 1)
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Error,
                    ErrorType = ErrorType.MissingUsedTimesetForPayload,
                    SheetName = sheetName,
                    RowNum = charRow.RowNum,
                    ColList = col,
                    Message = $"None of timeset for payload({charRow.MappingKey.Payload}) in base program.",
                    CommentsList = new List<string> { },
                });
                foreach (int column in col)
                {
                    ErrorReportManager.AddError(CharErrorType.E_MissingUsedTimesetForPayload_01, sheetName, charRow.RowNum, column,
                        [charRow.MappingKey.Payload],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
        }
        private void _CheckMappingDcCategory(Characterization charRow, string sheetName, List<int> col)
        {
            if (charRow.OtherSupplies.Split(' ').Length > 1)
            {
                if (!string.IsNullOrEmpty(charRow.OtherSupplies.Split(' ')[0]))
                {
                    return;
                }
            }
            else if (charRow.OtherSupplies.Split(' ').Length == 1)
            {
                if (string.IsNullOrEmpty(charRow.DcSelector))
                {
                    return; //When OtherSupplies does not contain H/L/NV, it is assumed that dc category definition exists.
                }
            }

            if (charRow.MappingSpec.DcCategoryLevel.Count > 1)
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Warning,
                    ErrorType = ErrorType.MultiUsedDcForPayload,
                    SheetName = sheetName,
                    RowNum = charRow.RowNum,
                    ColList = col,
                    Message =
                        $"Multi dc category for payload({charRow.MappingKey.Payload}) in base program: {string.Join(",", charRow.MappingSpec.DcCategoryLevel)} .",
                    CommentsList = new List<string> { },
                });
                foreach (int column in col)
                {
                    ErrorReportManager.AddError(CharErrorType.W_MultiUsedDcForPayload_01, sheetName, charRow.RowNum, column,
                        [charRow.MappingKey.Payload, string.Join(",", charRow.MappingSpec.DcCategoryLevel)],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
            if (charRow.MappingSpec.DcCategoryLevel.Count < 1)
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Error,
                    ErrorType = ErrorType.MissingUsedDcForPayload,
                    SheetName = sheetName,
                    RowNum = charRow.RowNum,
                    ColList = col,
                    Message = $"None of dc category for payload({charRow.MappingKey.Payload}) in base program.",
                    CommentsList = new List<string> { },
                });
                foreach (int column in col)
                {
                    ErrorReportManager.AddError(CharErrorType.E_MissingUsedDcForPayload_01, sheetName, charRow.RowNum, column,
                        [charRow.MappingKey.Payload],
                        new ErrorInfo() { Comments = new List<string> { } });
                }
            }
        }
        //private void _CheckMappingLevel(Characterization charRow, string sheetName, List<int> col)
        //{
        //    if (charRow.MappingSpec.Level.Count > 1)
        //    {
        //        ErrorMessages.Add(new ErrorMessage
        //        {
        //            ErrorLevel = ErrorLevel.Warning,
        //            ErrorType = ErrorType.MappingLevelMoreThanOne,
        //            SheetName = sheetName,
        //            RowNum = charRow.RowNum,
        //            ColList = col,
        //            Message = string.Format("Number of mapping level sheet by payload({0}) is more than 1, tool get level sheets: {1}", charRow.MappingKey.Payload, string.Join(",", charRow.MappingSpec.Level)),
        //            CommentsList = new List<string> { },
        //        });
        //    }
        //    if (charRow.MappingSpec.Level.Count < 1)
        //    {
        //        ErrorMessages.Add(new ErrorMessage
        //        {
        //            ErrorLevel = ErrorLevel.Error,
        //            ErrorType = ErrorType.MappingLevelLessThanOne,
        //            SheetName = sheetName,
        //            RowNum = charRow.RowNum,
        //            ColList = col,
        //            Message = string.Format("Number of mapping level sheet by payload({0}) is less than 1", charRow.MappingKey.Payload),
        //            CommentsList = new List<string> { },
        //        });
        //    }
        //}
    }
}
