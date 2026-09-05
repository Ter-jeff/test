using System.Collections.Generic;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class UslLslChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization charRow in charList)
            {
                if (charRow.IpUse3 != "")
                {
                    if (!double.TryParse(charRow.IpUse3, out _))
                    {
                        const string outString = "Can not convert USL to a value for limit!";
                        ErrorManager.AddError(ErrorType.IllegalCharUslLsl, charRow.SheetName, charRow.RowNum,
                            charRow.ColNum(new List<string> { "usl" }), charRow.Use, outString);
                        foreach (int col in charRow.ColNum(new List<string> { "usl" }))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_IllegalCharUslLsl_01, charRow.SheetName, charRow.RowNum, col, ["USL"],
                                new ErrorInfo() { Comments = new List<string> { "usl" } });
                        }
                    }
                }

                if (charRow.IpUse4 != "")
                {
                    if (!double.TryParse(charRow.IpUse4, out _))
                    {
                        const string outString = "Can not convert LSL to a value for limit!";
                        ErrorManager.AddError(ErrorType.IllegalCharUslLsl, charRow.SheetName, charRow.RowNum,
                            charRow.ColNum(new List<string> { "lsl" }), charRow.Use, outString);
                        foreach (int col in charRow.ColNum(new List<string> { "lsl" }))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_IllegalCharUslLsl_01, charRow.SheetName, charRow.RowNum, col, ["LSL"],
                                new ErrorInfo() { Comments = new List<string> { "lsl" } });
                        }
                    }
                }

                if (charRow.IpUse3 != "" && charRow.IpUse4 != "")
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
}
