using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow
{
    public class FreqFlowRowGenerator : HardIpFlowRowGenerator
    {
        public FreqFlowRowGenerator(HardIpInputData hardIpInputData, string sheetName)
            : base(hardIpInputData, sheetName)
        {
        }

        protected internal override string CreateTestFailAction() => CreateGenericFailAction();

        protected override string CreateUseLimitFailAction(bool forShmoo) => CreateGenericFailAction();

        private string CreateGenericFailAction()
        {
            string patternName = Pattern.Pattern.GetPatternName();
            if (Pattern.Pattern.IsMultiple())
            {
                patternName = patternName.Replace(PatternClass.MultipleConst, "");
                string failFlag = CommonGenerator.GenHardIpFlowFailAction(
                    SheetName,
                    BlockName,
                    SubBlockName.ToUpper().Replace("-MERGE", ""),
                    patternName,
                    TimingAc,
                    InstNameSubStr,
                    ActualLabelVoltage,
                    Pat.MiscInfo,
                    NoPattern
                );
                ErrorReportManager.AddError(ClockCheckErrorType.I_BurstPatJudgeFlag_01, Pattern.SheetName, Pattern.RowNum, 1, [failFlag]);
                return failFlag;
            }

            return CommonGenerator.GenHardIpFlowFailAction(
                SheetName,
                BlockName,
                SubBlockName.ToUpper().Replace("-MERGE", ""),
                patternName,
                TimingAc,
                InstNameSubStr,
                ActualLabelVoltage,
                Pat.MiscInfo,
                NoPattern
            );
        }
    }
}
