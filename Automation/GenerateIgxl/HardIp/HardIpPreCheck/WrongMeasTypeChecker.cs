using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class WrongMeasTypeChecker : HardIpPrecheckBase
    {
        public WrongMeasTypeChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        public override void Check()
        {
            int measIndex = HardIpSheet.PlanHeaderIdx["measIndex"];
            foreach (MeasPin pin in Pattern.MeasPins)
            {
                if (!MeasType.MeasTypes.Contains(pin.MeasType))
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongMeasType_01,
                        Pattern.SheetName,
                        pin.RowNum,
                        measIndex,
                        [pin.MeasType]
                    );
                }
            }
        }
    }
}
