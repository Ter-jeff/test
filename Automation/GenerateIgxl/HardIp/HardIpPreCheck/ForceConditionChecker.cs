using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Singleton;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class ForceConditionChecker : HardIpPrecheckBase
    {
        private readonly int _forceIndex;
        private readonly List<string> _selectorVoltage = new List<string> { HardIpConstData.LabelAll, HardIpConstData.LabelNv, HardIpConstData.LabelLv, HardIpConstData.LabelHv };
        private readonly List<string> _selectorType = new List<string> { HardIpConstData.SelectTyp, HardIpConstData.SelectMax, HardIpConstData.SelectMin };
        private readonly List<string> _forceTypes = new List<string> { "VT", "VIH", "VIL", "VOL", "VOH", "V", "I", "IOH", "IOL",
            "VID", "VICM", "VCM", "TERM", "V1P", "V2P", "V1N", "V2N", "VOD", "DISABLE_FRC","DIFFVT" ,"VALT"};
        private readonly List<string> _forceMethods = new List<string> {
            "CONNECTPPMU", "DISCONNECTPPMU", "CONNECTDIGITAL", "DISCONNECTDIGITAL", "DISABLECOMPARE", "CONNECTDCVI", "DISCONNECTDCVI",
            "RELAYON", "RELAYOFF", "RELAY_ON", "RELAY_OFF", "INITHI", "INITLO", "INITHIZ", "DISABLE_FRC" };

        public ForceConditionChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
            _forceIndex = hardIpSheet.ForceIndex;
        }

        public static bool CheckAcSettingPin(HardIpPattern pattern, out string wrongPin)
        {
            wrongPin = "";
            string timeSetUsedStr = Regex.Replace(pattern.TimeSetUsed.McgSetting, "::", "&");
            if (!AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(pattern.Pattern.GetLastPayload()))
            {
                return false;
            }


            string timeSetName = AcTSetCategoryMapSingleton.Instance().PatternList[pattern.Pattern.GetLastPayload()].TimeSetVersion;
            TimeSetBasicSheet timeSet = TestProgram.IgxlWorkBk.TimeSetSheets.Values.FirstOrDefault(s => s.Name.Equals(timeSetName, StringComparison.OrdinalIgnoreCase));
            if (timeSet == null)
            {
                wrongPin = "";
                return false;
            }

            foreach (string changedPin in timeSetUsedStr.Trim(';').Split(';'))
            {
                string pinName = changedPin.Split(':')[1];
                bool pinExist = false;
                foreach (TSet timeSetData in timeSet.Rows)
                {
                    foreach (TimingRow timeSetRow in timeSetData.TimingRows)
                    {
                        if (timeSetRow.PinGrpName.Equals(pinName, StringComparison.OrdinalIgnoreCase))
                        {
                            pinExist = true;
                            break;
                        }
                        if (pinName.Split('&').ToList().Exists(s => s.Equals(timeSetRow.PinGrpName, StringComparison.OrdinalIgnoreCase)))
                        {
                            pinExist = true;
                            break;
                        }
                    }
                }

                if (!pinExist)
                {
                    wrongPin += Regex.Replace(pinName, "&", "::") + ",";
                }
            }
            wrongPin = wrongPin.Trim(',');

            if (!string.IsNullOrEmpty(wrongPin))
            {
                return false;
            }

            return true;
        }

        internal void CheckForcePin(HardIpPattern pattern, ForcePin forcePin, MeasPin measpin = null)
        {
            int rowNum;
            if (measpin != null)
            {
                rowNum = measpin.RowNum;
            }
            else
            {
                rowNum = pattern.RowNum;
            }

            if (forcePin.Type == ForceConditionType.Normal && !_forceTypes.Contains(forcePin.ForceType.ToUpper()))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongForceCondition_01,
                    pattern.SheetName,
                    rowNum,
                    _forceIndex,
                    [forcePin.ForceType]
                );
            }
            else if (forcePin.Type == ForceConditionType.Others && !_forceMethods.Contains(forcePin.ForceValue.ToUpper()))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongForceCondition_02,
                    pattern.SheetName,
                    rowNum,
                    _forceIndex,
                    [forcePin.ForceValue]
                );
            }
        }

        internal void CheckRelay(HardIpPattern pattern, ForcePin forcePin)
        {
            int rowNum = pattern.RowNum;
            if (forcePin.Type == ForceConditionType.Others && Regex.IsMatch(forcePin.ForceValue, "relay_*On|Off", RegexOptions.IgnoreCase))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.W_RelayRestore_01,
                    pattern.SheetName,
                    rowNum,
                    _forceIndex,
                    [forcePin.PinName]
                );
            }
        }

        public override void Check()
        {
            #region Check illegal syntax for pattern's force condition
            if (Pattern.ForceConditionList.Count > 0)
            {
                ForceCondition condition = Pattern.ForceConditionList[0];
                foreach (ForcePin forcePin in condition.ForcePins)
                {
                    CheckForcePin(Pattern, forcePin);
                    CheckRelay(Pattern, forcePin);
                }
            }
            #endregion

            #region Check illegal syntax for meaPin's force condition
            foreach (MeasPin pin in Pattern.MeasPins)
            {
                CheckMeasPinForceCondition(pin);
            }
            #endregion

            #region Check wrong AC setting pin, Ac setting pin should be "TCK", "ShiftIn", Nwire pin or pin in timeset sheet pattern used
            if (!string.IsNullOrEmpty(Pattern.TimeSetUsed.McgSetting))
            {
                if (!CheckAcSettingPin(Pattern, out string wrongPin) && wrongPin != "")
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongForceCondition_05,
                        Pattern.SheetName,
                        Pattern.RowNum,
                        _forceIndex,
                        [wrongPin]
                    );
                }
            }
            #endregion

            #region Check ADSelector and DCSelector
            if (!string.IsNullOrEmpty(Pattern.AcSelectorUsed))
            {
                CheckAcSelector();
            }
            #endregion
        }

        internal void CheckMeasPinForceCondition(MeasPin pin)
        {
            if (pin.ForceConditions.Count == 0)
            {
                return;
            }

            #region check illegal force type for measPin
            ForceCondition condition = pin.ForceConditions[0];
            foreach (ForcePin forcePin in condition.ForcePins)
            {
                CheckForcePin(Pattern, forcePin, pin);
            }
            #endregion

            #region

            if (pin.MeasType == "" || pin.MeasType.Equals(MeasType.MeasC) ||
                pin.MeasType.Equals(MeasType.MeasCalc))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongForceCondition_03,
                    Pattern.SheetName,
                    pin.RowNum,
                    _forceIndex,
                    [pin.MeasType]
                );
            }

            #endregion


            #region Check misMatched force type
            if (SearchInfo.GetPinType(pin.PinName) == "I/O")
            {
                if ((pin.MeasType.ToUpper() == "MEASV" && condition.ForcePins.Where(x => x.PinName == pin.PinName).ToList().Exists(y => y.ForceType.ToUpper() == "V"))
                    || (pin.MeasType.ToUpper() == "MEASI" && condition.ForcePins.Where(x => x.PinName == pin.PinName).ToList().Exists(y => y.ForceType.ToUpper() == "I")))
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongForceCondition_04,
                        Pattern.SheetName,
                        pin.RowNum,
                        _forceIndex,
                        [pin.MeasType]
                    );
                }
            }
            #endregion
        }

        internal void CheckAcSelector()
        {
            foreach (string info in Pattern.AcSelectorUsed.Split(';'))
            {
                string[] acArray = info.Split(':');
                if (acArray.Length == 3)
                {
                    if (!_selectorVoltage.Contains(acArray[1], StringComparer.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_WrongForceCondition_06,
                            Pattern.SheetName,
                            Pattern.RowNum,
                            _forceIndex,
                            [acArray[1]]
                        );
                    }
                    if (!_selectorType.Contains(acArray[2], StringComparer.OrdinalIgnoreCase))
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_WrongForceCondition_07,
                            Pattern.SheetName,
                            Pattern.RowNum,
                            _forceIndex,
                            [acArray[2]]
                        );
                    }
                }
                else
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongForceCondition_08,
                        Pattern.SheetName,
                        Pattern.RowNum,
                        _forceIndex,
                        [info]
                    );
                }
            }
        }
    }
}
