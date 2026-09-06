using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class DuplicateInstanceChecker : HardIpPrecheckBase
    {
        public DuplicateInstanceChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        public override void Check()
        {
            if (Pattern.ForceConditionList.Count > 1)
            {
                return;
            }

            if (Pattern.DupIndex > 0)
            {
                string errorMessage = "We found duplicate instance but we add index to avoid validate fail, please check";
                Response.Report(errorMessage, EnumMessageLevel.Warning, 0);
                ErrorReportManager.AddError(
                    HardIpErrorType.W_DuplicateInstance_01,
                    Pattern.SheetName,
                    Pattern.RowNum,
                    0,
                    []
                );
            }
        }
    }
}
