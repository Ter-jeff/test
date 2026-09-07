using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class OppositeUslLslChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization charRow in charList
                .Where(charRow => charRow.IpUse3 != "" && charRow.IpUse4 != ""))
            {
                double.TryParse(charRow.IpUse3, out double usl);
                double.TryParse(charRow.IpUse4, out double lsl);

                if (!(usl < lsl))
                {
                    continue;
                }

                const string outString = "Opposite USL/LSL ";
                ErrorManager.AddError(ErrorType.OppositeUsl, charRow.SheetName, charRow.RowNum,
                    charRow.ColNum(new List<string> { "usl", "lsl" }), charRow.Use, outString);

                foreach (int col in charRow.ColNum(new List<string> { "usl", "lsl" }))
                {
                    ErrorReportManager.AddError(CharErrorType.E_OppositeUsl_01, charRow.SheetName, charRow.RowNum, col, []);
                }
            }
        }
    }
}
