using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class MissingPinNameChecker : HardIpPrecheckBase
    {
        public MissingPinNameChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        public override void Check()
        {
            int measIndex = HardIpSheet.PlanHeaderIdx["measIndex"];
            foreach (MeasPin pin in Pattern.MeasPins)
            {
                string limits = pin.MeasLimitsH.Aggregate("", (current, limit) => current + limit.LoLimit + limit.HiLimit);
                limits += pin.MeasLimitsL.Aggregate(limits, (current, limit) => current + limit.LoLimit + limit.HiLimit);
                limits += pin.MeasLimitsN.Aggregate(limits, (current, limit) => current + limit.LoLimit + limit.HiLimit);
                if (pin.PinName == "" && limits != "")
                {
                    //const string errorMessage = "Empty Meas PinName but contains Limits";
                    ErrorReportManager.AddError(HardIpErrorType.E_MissingPinName_07, Pattern.SheetName, pin.RowNum, measIndex);
                }
            }
        }
    }
}
