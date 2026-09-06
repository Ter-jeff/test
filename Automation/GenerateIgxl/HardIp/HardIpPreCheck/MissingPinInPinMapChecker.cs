using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class MissingPinInPinMapChecker : HardIpPrecheckBase
    {
        private int _errorCnt;

        public MissingPinInPinMapChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        private void CheckOneLimit(string sheetName, MeasPin pin, MeasLimit limit)
        {
            if (Regex.IsMatch(limit.HiLimit, @"(?<pinName>(VDD\w+))", RegexOptions.IgnoreCase))
            {
                string pinName = Regex.Match(limit.HiLimit, @"(?<pinName>(VDD\w+))", RegexOptions.IgnoreCase).Groups["pinName"].ToString().Trim().ToUpper();
                if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(pinName) && _errorCnt < 500)
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_MissingPinName_01, sheetName, pin.RowNum, limit.HiHeaderIndex, [pinName]);
                    _errorCnt++;
                }
            }

            if (Regex.IsMatch(limit.LoLimit, @"(?<pinName>(VDD\w+))", RegexOptions.IgnoreCase))
            {
                string pinName = Regex.Match(limit.LoLimit, @"(?<pinName>(VDD\w+))", RegexOptions.IgnoreCase).Groups["pinName"].ToString().Trim().ToUpper();
                if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(pinName) && _errorCnt < 500)
                {
                    ErrorReportManager.AddError(HardIpErrorType.E_MissingPinName_02, sheetName, pin.RowNum, limit.LoHeaderIndex, [pinName]);
                    _errorCnt++;
                }
            }
        }

        public override void Check()
        {
            int forceIndex = HardIpSheet.PlanHeaderIdx["forceIndex"];
            int measIndex = HardIpSheet.PlanHeaderIdx["measIndex"];

            CheckPatternForcePins(forceIndex);
            CheckSweepPins(forceIndex);

            foreach (MeasPin pin in Pattern.MeasPins)
            {
                CheckMeasPin(pin, forceIndex, measIndex);
            }
        }

        private void CheckPatternForcePins(int forceIndex)
        {
            #region Check force pins for pattern
            foreach (ForceCondition condition in Pattern.ForceConditionList)
            {
                foreach (ForcePin forcePin in condition.ForcePins)
                {
                    string forcePinName = forcePin.PinName.ToUpper();
                    if (forcePinName.Contains("::"))
                    {
                        forcePinName = forcePinName.Split(':')[0];
                    }
                    else if (forcePinName.Contains("="))
                    {
                        forcePinName = forcePinName.Split('=')[1];
                    }

                    if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(forcePinName.ToUpper()) && _errorCnt < 500)
                    {
                        ErrorReportManager.AddError(HardIpErrorType.E_MissingPinName_03, Pattern.SheetName, Pattern.RowNum, forceIndex, [forcePinName]);
                        _errorCnt++;
                    }
                }
            }
            #endregion
        }

        private void CheckSweepPins(int forceIndex)
        {
            #region Check sweep pins for pattern
            foreach (KeyValuePair<string, List<SweepVData>> sweepVoltage in Pattern.SweepVoltage)
            {
                foreach (SweepVData sweep in sweepVoltage.Value)
                {
                    sweep.PinName.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList().ForEach
                    (pinmame =>
                    {
                        if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(pinmame.ToUpper()) && _errorCnt < 500)
                        {
                            ErrorReportManager.AddError(HardIpErrorType.E_MissingPinName_04, Pattern.SheetName, Pattern.RowNum, forceIndex, [pinmame]);
                            _errorCnt++;
                        }
                    });
                }
            }
            #endregion
        }

        private void CheckMeasPin(MeasPin pin, int forceIndex, int measIndex)
        {
            var missingPinList = new List<string>();

            #region Check MeasPins
            string measPinName = pin.PinName;
            foreach (string subPin in measPinName.Split(','))
            {
                string tmpPin = subPin;
                if (subPin.Contains("::"))
                {
                    tmpPin = subPin.Split(':')[0];
                }
                else if (subPin.Contains("="))
                {
                    tmpPin = subPin.Split('=')[1];
                }

                if (tmpPin != "" && !(pin.MeasType == MeasType.MeasCalc || pin.MeasType == MeasType.MeasLimit))
                {
                    if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(tmpPin))
                    {
                        missingPinList.Add(subPin);
                    }
                }
            }

            if (missingPinList.Count > 0)
            {
                ErrorReportManager.AddError(HardIpErrorType.E_MissingPinName_05, Pattern.SheetName, pin.RowNum, measIndex, missingPinList.ToArray());
                _errorCnt++;
            }
            #endregion

            #region Check force pins of MeasPin
            foreach (ForceCondition condition in pin.ForceConditions)
            {
                foreach (ForcePin forcePin in condition.ForcePins)
                {
                    measPinName = forcePin.PinName.ToUpper();
                    if (measPinName.Contains("::"))
                    {
                        measPinName = measPinName.Split(':')[0];
                    }
                    else if (measPinName.Contains("="))
                    {
                        measPinName = measPinName.Split('=')[1];
                    }

                    if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(measPinName))
                    {
                        if (TestProgram.IgxlWorkBk.PinMapPair.Value != null && TestProgram.IgxlWorkBk.PinMapPair.Value.GroupList.All(x => x.PinName.ToUpper() != measPinName))
                        {
                            ErrorReportManager.AddError(HardIpErrorType.E_MissingPinName_06, Pattern.SheetName, pin.RowNum, forceIndex, [forcePin.PinName]);
                            _errorCnt++;
                        }
                    }
                }
            }
            #endregion

            #region Check Spec pins used in limit value equation
            foreach (MeasLimit limit in pin.MeasLimitsH)
            {
                CheckOneLimit(Pattern.SheetName, pin, limit);
            }
            foreach (MeasLimit limit in pin.MeasLimitsL)
            {
                CheckOneLimit(Pattern.SheetName, pin, limit);
            }
            foreach (MeasLimit limit in pin.MeasLimitsN)
            {
                CheckOneLimit(Pattern.SheetName, pin, limit);
            }

            #endregion
        }
    }
}
