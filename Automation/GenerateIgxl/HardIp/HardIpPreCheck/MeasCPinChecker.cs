using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class MeasCPinChecker : HardIpPrecheckBase
    {
        public MeasCPinChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        public override void Check()
        {
            if (HardIpConstData.RegInsInPatt.IsMatch(Pattern.Pattern.RealPatternName) ||
                HardIpConstData.RegOpcodeInPatt.IsMatch(Pattern.Pattern.RealPatternName))
            {
                return;
            }

            int measIdx = HardIpSheet.PlanHeaderIdx["measIndex"];
            CheckPinWithInfo(Pattern, measIdx);
        }

        private void CheckPinWithInfo(HardIpPattern pattern, int measIdx)
        {
            HardIpInfo patInfo = HardIpService.GetHardIpInfo(pattern);
            string infoPin = patInfo.CapPinName;
            IEnumerable<MeasPin> measCRow = pattern.MeasPins.Where(x => x.MeasType == "MeasC");

            if (!measCRow.Any())
            {
                return;
            }

            if (string.IsNullOrEmpty(infoPin) && !pattern.Pattern.IsMultiTimeDomain())
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongMeasC_01,
                    pattern.SheetName,
                    pattern.RowNum,
                    measIdx,
                    [patInfo.Payload]
                );
            }
            else
            {
                if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(infoPin.ToUpper()))
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_MissingPatInfoPinInPinMap_01,
                        pattern.SheetName,
                        pattern.RowNum,
                        measIdx,
                        [infoPin]
                    );
                    return;
                }
                foreach (MeasPin measC in measCRow)
                {
                    if (!measC.PinName.Equals(infoPin, StringComparison.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_WrongMeasPinInPatInfo_01,
                            pattern.SheetName,
                            measC.RowNum,
                            measIdx,
                            [measC.PinName, infoPin]
                        );
                    }
                }
            }
        }
    }
}
