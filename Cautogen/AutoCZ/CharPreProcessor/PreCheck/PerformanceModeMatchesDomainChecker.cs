using System.Collections.Generic;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    internal class PerformanceModeMatchesDomainChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charRows, string sheetName)
        {
            /* check if group (performance mode) of charRow matches USERDEF2 (domain) */
            foreach (Characterization charRow in charRows)
            {
                // bypass check group for USERDEF1 == MCL/H
                if (charRow.UserDef1.ToUpper() == "MCL" || charRow.UserDef1.ToUpper() == "MCH")
                {
                    return;
                }

                string outString = "";
                string group = "";
                switch (charRow.UserDef2.ToUpper())
                {
                    case "CPU":
                        if (charRow.Group == "" || charRow.Group.ToUpper().StartsWith("MS") || charRow.Group.ToUpper().StartsWith("MG"))
                        {
                            outString = "Wrong group for CPU";
                            group = "CPU";
                        }

                        break;

                    case "GPU":
                        if (charRow.Group == "" || charRow.Group.ToUpper().StartsWith("MS") || charRow.Group.ToUpper().StartsWith("MC"))
                        {
                            outString = "Wrong group for GPU";
                            group = "GPU";
                        }

                        break;

                    case "SOC":
                        if (charRow.Group == "" || charRow.Group.ToUpper().StartsWith("MC") || charRow.Group.ToUpper().StartsWith("MG"))
                        {
                            outString = "Wrong group for SOC";
                            group = "SOC";
                        }

                        break;
                }
                if (outString == "")
                {
                    continue;
                }

                // report warning
                ErrorManager.AddWarning(ErrorType.WrongGroup, charRow.SheetName, charRow.RowNum, charRow.ColNum("group"), charRow.Use, outString);
                foreach (int col in charRow.ColNum("group"))
                {
                    ErrorReportManager.AddError(CharErrorType.W_WrongGroup_01, charRow.SheetName, charRow.RowNum, col, [group]);
                }
            }
        }
    }
}
