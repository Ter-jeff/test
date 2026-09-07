using System;
using System.Collections.Generic;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class UslLslRangeChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization item in charList)
            {
                if (item.Pins.Count <= 0 || (item.IpUse3 == "" && item.IpUse4 == "") ||
                    (item.IpUse3 + item.IpUse4).Contains(","))
                {
                    continue;
                }

                string start = "";
                string stop = "";
                string message = "";
                foreach (Pin pin in item.Pins)
                {
                    start = pin.Start;
                    stop = pin.Stop;
                    if (start != stop)
                    {
                        break;
                    }
                }

                if (start == "" || stop == "")
                {
                    continue;
                }

                double myLimit;
                foreach (string usl in item.IpUse3.Split(':'))
                {
                    if (usl != "")
                    {
                        if (double.TryParse(usl.Trim(), out myLimit))
                        {
                            if ((Convert.ToDouble(start) - myLimit) * (Convert.ToDouble(stop) - myLimit) > 0)
                            {
                                message = "Wrong USL for shmoo range in " + item.SheetName;
                            }
                        }
                        else
                        {
                            message = "USL value is invalid in " + item.SheetName;
                        }
                    }
                    if (message == "")
                    {
                        continue;
                    }

                    string outString = message;
                    ErrorManager.AddError(ErrorType.WrongUsllslRange, item.SheetName, item.RowNum, item.ColNum("usl"), item.Use, outString);
                    foreach (int col in item.ColNum("usl"))
                    {
                        ErrorReportManager.AddError(CharErrorType.E_WrongUsllslRange_01, item.SheetName, item.RowNum, col, ["USL", item.SheetName]);
                    }
                }

                message = "";

                foreach (string lsl in item.IpUse4.Split(':'))
                {
                    if (lsl != "")
                    {
                        if (double.TryParse(lsl.Trim(), out myLimit))
                        {
                            if ((Convert.ToDouble(start) - myLimit) * (Convert.ToDouble(stop) - myLimit) > 0)
                            {
                                message = "Wrong LSL for shmoo range in " + item.SheetName;
                            }
                        }
                        else
                        {
                            message = "LSL value is invalid in " + item.SheetName;
                        }
                    }
                    if (message == "")
                    {
                        continue;
                    }

                    string outString = message;
                    ErrorManager.AddError(ErrorType.WrongUsllslRange, item.SheetName, item.RowNum, item.ColNum("lsl"), item.Use, outString);
                    foreach (int col in item.ColNum("lsl"))
                    {
                        ErrorReportManager.AddError(CharErrorType.E_WrongUsllslRange_01, item.SheetName, item.RowNum, col, ["LSL", item.SheetName]);
                    }
                }
            }
        }
    }
}
