using System.Collections.Generic;
using System.Linq;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.InputChecker
{
    public class EfuseConfigureMainChecker
    {
        public void WorkFlow(EfuseConfigMainSheet sheet)
        {
            bool needSort = sheet.Rows.First().NeedSort;
            foreach (EfuseConfigMainRow row in sheet.Rows)
            {
                CheckIfLocationIsFtf(sheet, row);
                CheckFuseColumnFormat(sheet, row, needSort);
            }
        }

        private void CheckIfLocationIsFtf(EfuseConfigMainSheet sheet, EfuseConfigMainRow row)
        {
            if (row.FuseBlowLocation.ContainsIgnoreCase("FTF"))
            {
                ErrorReportManager.AddError(EFuseErrorType.W_RuleViolationFuseBlowLocation_01, sheet.SheetName, row.RowNum, sheet.FuseBlowLocationColNumber, []);
            }
        }

        private void CheckFuseColumnFormat(EfuseConfigMainSheet sheet, EfuseConfigMainRow row, bool firstNeedSort)
        {
            if (sheet.FuseColNumber != -1)
            {
                string lsbMsb = row.Fuse.Replace("[", "");
                lsbMsb = lsbMsb.Replace("]", "");
                List<string> lsbMsbStr = lsbMsb.Split(':').ToList();
                //string message = "";
                if (lsbMsbStr.Count == 1)
                {
                    if (!lsbMsbStr.First().IsNumber())
                    {
                        //message = $"The content of this column must be [MSB:LSB] or [LSB:MSB]. Content {row.Fuse}";
                        ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_07, sheet.SheetName, row.RowNum, sheet.FuseColNumber, [row.Fuse]);
                    }
                }
                else if (lsbMsbStr.Count == 2)
                {
                    if (!(lsbMsbStr[0].IsNumber() && lsbMsbStr[1].IsNumber()))
                    {
                        //message = $"Non-numeric string exists in Fuse column. Content {row.Fuse}";
                        ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_08, sheet.SheetName, row.RowNum, sheet.FuseColNumber, [row.Fuse]);
                    }
                }

                if (firstNeedSort != row.NeedSort)
                {
                    //string msg;
                    if (firstNeedSort)
                    {
                        //msg = $"This row format is MSB:LSB is inconsistent with the LSB:MSB used by other rows. Content {row.Fuse}";
                        ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_09, sheet.SheetName, row.RowNum, sheet.FuseColNumber, [row.Fuse]);
                    }
                    else
                    {
                        //msg = $"This row format is LSB:MSB is inconsistent with the MSB:LSB used by other rows. Content {row.Fuse}";
                        ErrorReportManager.AddError(EFuseErrorType.E_RuleViolationBit_10, sheet.SheetName, row.RowNum, sheet.FuseColNumber, [row.Fuse]);
                    }
                }

            }
        }
    }
}
