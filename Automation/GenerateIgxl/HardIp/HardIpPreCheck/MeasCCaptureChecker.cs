using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class MeasCCaptureChecker : HardIpPrecheckBase
    {
        public MeasCCaptureChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        public override void Check()
        {
            int measIndex = HardIpSheet.PlanHeaderIdx["measIndex"];
            var temp = Pattern.MeasPins.Where(x => !string.IsNullOrEmpty(x.CusStr)).ToList();
            var groups = temp.GroupBy(p => p.SequenceIndex).ToDictionary(p => p.Key, p => p.ToList());

            foreach (KeyValuePair<int, List<MeasPin>> item in groups)
            {
                int count = 0;
                var dic = new Dictionary<string, string>();
                foreach (MeasPin measPin in item.Value)
                {
                    if (dic.ContainsKey(measPin.CusStr))
                    {
                        ++count;
                        string errorMessage =
                            $"The capture store name : {measPin.CusStr} is repeated with {count} times in this item! ";
                        ErrorReportManager.AddError(
                            HardIpErrorType.W_WrongMeasC_01,
                            Pattern.SheetName,
                            measPin.RowNum,
                            measIndex,
                            [measPin.CusStr, $"{count}"]
                        );
                    }
                    else
                    {
                        dic.Add(measPin.CusStr, measPin.CusStr);
                    }
                }
            }


        }
    }
}
