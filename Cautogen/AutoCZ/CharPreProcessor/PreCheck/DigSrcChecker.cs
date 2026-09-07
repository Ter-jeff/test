using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager.PatternError;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class DigSrcChecker : PreCheckBase
    {
        public static Regex RexDigSrc = new Regex(":DigSrc", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public override void Check(List<Characterization> charList, string sheetName)
        {
            if (UtilityMain.UtilityData.PatInfoDict.Count == 0)
            {
                ErrorMessages.Add(new ErrorMessage
                {
                    ErrorLevel = ErrorLevel.Warning,
                    ErrorType = ErrorType.MissingPatInfo,
                    Message = "Pattern info is empty, please check!",
                });
                ErrorReportManager.AddError(CharErrorType.W_MissingPatInfo_01, "", 0, 0, []);
                return;
            }

            foreach (Characterization item in charList.Where(IsUseItem))
            {
                //if rtos cmd, do not check pattern information
                if (item.IsUseRtosCmd)
                {
                    continue;
                }

                item.PatternCellList.ForEach(x => _Check(x, sheetName, item.RowNum));
            }
        }

        /* pattern:DigSrc need to have sendBitStr in pattern info */
        private void _Check(PatternCell patternCell, string sheetName, int rowNum)
        {
            IEnumerable<string> splitPatterns = patternCell.PatternDefine.Split(',').Select(x => x.Trim());
            foreach (string splitPattern in splitPatterns)
            {
                //Pass empty string
                if (string.IsNullOrEmpty(splitPattern))
                {
                    continue;
                }

                //Pass non Pat:DigSrc and non SRMDSSC
                var splitDssc = splitPattern.Split(':').Select(x => x.Trim()).ToList();
                if (splitDssc.Count < 2 && !splitPattern.ToUpper().Contains("SRMDSSC"))
                {
                    continue;
                }

                string pattern = splitDssc[0];

                // check pattern is in pat info
                if (!UtilityMain.UtilityData.PatInfoDict.ContainsKey(pattern.ToUpper()))
                {
                    PatternErrorCache.Mark(pattern, sheetName, PatternErrorCache.MarkType.MissingPatternInPatInfo);
                    ErrorMessages.Add(new ErrorMessage
                    {
                        ErrorLevel = ErrorLevel.Error,
                        ErrorType = ErrorType.MisMatchDigSrc,
                        SheetName = sheetName,
                        RowNum = rowNum,
                        ColList = new List<int>() { patternCell.ColIndex },
                        Message = "Pattern ends with DigSrc without SendBitStr in pat info!",
                        CommentsList = new List<string> { pattern + "(missing pattern info)" },
                    });
                    ErrorReportManager.AddError(CharErrorType.E_MisMatchDigSrc_01, sheetName, rowNum, patternCell.ColIndex, [],
                        new ErrorInfo() { Comments = new List<string>() { $"{pattern} (missing pattern info)" } });
                }
                else
                {
                    // check pattern in pattern info contains SendBitStr
                    HardIpReference patInfo = UtilityMain.UtilityData.PatInfoDict[pattern.ToUpper()];
                    if (string.IsNullOrEmpty(patInfo.SendBitStr))
                    {
                        PatternErrorCache.Mark(pattern, sheetName, PatternErrorCache.MarkType.MissingPatternInPatInfo);
                        ErrorMessages.Add(new ErrorMessage
                        {
                            ErrorLevel = ErrorLevel.Error,
                            ErrorType = ErrorType.MisMatchDigSrc,
                            SheetName = sheetName,
                            RowNum = rowNum,
                            ColList = new List<int>() { patternCell.ColIndex },
                            Message = "Pattern ends with DigSrc without SendBitStr in pat info!",
                            CommentsList = new List<string> { pattern + "(missing sendBitStr in pattern info)" },
                        });
                        ErrorReportManager.AddError(CharErrorType.E_MisMatchDigSrc_01, sheetName, rowNum, patternCell.ColIndex, [],
                        new ErrorInfo() { Comments = new List<string>() { $"{pattern} (missing sendBitStr in pattern info)" } });
                    }
                }
            }
        }
    }
}
