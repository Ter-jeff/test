using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class MixedSiDmChecker : PreCheckBase
    {
        /* Check Mixed SI/DM */
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization item in charList)
            {
                var allPatterns = item.PatternCellList.SelectMany(x => x.PatternDefine.Split(',')).ToList();
                if (!allPatterns.Any(x => CheckEndOrContain(x, "SI")) || !allPatterns.Any(x => CheckEndOrContain(x, "DM")))
                {
                    continue;
                }

                var errPattern = new List<string>();
                foreach (PatternCell patternCell in item.PatternCellList)
                {
                    if (CheckEndOrContain(patternCell.PatternDefine, "SI") || CheckEndOrContain(patternCell.PatternDefine, "DM"))
                    {
                        errPattern.Add(patternCell.Header);
                    }
                }

                string outString = "init/payload patterns is mixed with SI/DM in " + item.SheetName + " Row " + item.RowNum;
                ErrorManager.AddError(ErrorType.MixedSIandDm, item.SheetName, item.RowNum, item.ColNum(errPattern), item.Use, outString);
                foreach (int col in item.ColNum(errPattern))
                {
                    ErrorReportManager.AddError(CharErrorType.E_MixedSIandDm_01, item.SheetName, item.RowNum, col, [item.SheetName, $"{item.RowNum}"]);
                }
            }
        }
        private bool CheckEndOrContain(string str, string comp)
        {
            str = str.ToUpper();
            return str.Contains(string.Format($"_{comp.ToUpper()}_"))
                || str.Contains(string.Format($"_{comp.ToUpper()}:"))
                || str.EndsWith(string.Format($"_{comp.ToUpper()}"));
        }
    }
}
