using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class ManualItemsChecker : HardIpPrecheckBase
    {
        public ManualItemsChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        public override void Check()
        {
            int miscInfoIndex = HardIpSheet.PlanHeaderIdx["miscInfoIndex"];
            if (Regex.IsMatch(Pattern.MiscInfo, HardIpConstData.Manual, RegexOptions.IgnoreCase))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.W_ManualItems_01,
                    Pattern.SheetName,
                    Pattern.RowNum,
                    miscInfoIndex,
                    []
                );
            }
        }
    }
}
