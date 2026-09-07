using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Basic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.Basic;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class TimeSetSheetGenerator
    {
        private const string Mcg = "_MCG";
        private const string Clock = "clock";
        private const string Clock2X = "clock_2X";
        private Dictionary<string, Dictionary<string, string>> _duplicateTimeSets = new Dictionary<string, Dictionary<string, string>>();
        private List<TimeSetBasicSheet> _timeSetSheets = new List<TimeSetBasicSheet>();
        private static readonly Regex _regex = new Regex(@"^[a-z|A-Z|_][\w]*", RegexOptions.IgnoreCase);
        private static readonly Regex _regex2 = new Regex("::");

        internal string ConvertUnits(string timeUsedStr)
        {
            var pins = new List<string>();
            foreach (string changedPin in timeUsedStr.Trim(';').Split(';'))
            {
                string pinName = changedPin.Split(':')[1];
                string clockValue = HardIpPattern.GetFreq(changedPin.Split(':')[2]);
                pins.Add(pinName + ":" + clockValue);
            }
            pins.Sort(string.CompareOrdinal);
            return string.Join(";", pins);
        }

        public List<TimeSetBasicSheet> GenTimeSet(Dictionary<string, HardIpSheet> planDic)
        {
            _duplicateTimeSets = new Dictionary<string, Dictionary<string, string>>();
            _timeSetSheets = new List<TimeSetBasicSheet>();
            foreach (HardIpSheet sheet in planDic.Values)
            {
                foreach (HardIpPattern pattern in sheet.Rows)
                {
                    if (string.IsNullOrEmpty(pattern.TimeSetUsed.McgSetting))
                    {
                        continue;
                    }
                    string lastPayload = pattern.Pattern.GetLastPayload();
                    if (AcTSetCategoryMapSingleton.Instance().PatternList.TryGetValue(lastPayload, out PatternData patternData))
                    {
                        string timeSetName = patternData.TimeSetVersion;
                        TimeSetBasicSheet timeSet = TestProgram.IgxlWorkBk.TimeSetSheets.Values.FirstOrDefault(s => s.Name.Equals(timeSetName, StringComparison.OrdinalIgnoreCase));
                        if (timeSet == null)
                        {
                            continue;
                        }

                        UpdateTimeSetSheetWithMcg(timeSet, pattern);
                    }
                }
            }
            return _timeSetSheets;
        }

        internal string GetPinClockPeriod(string pinName, string newValue, TimeSetBasicSheet timeSetSheet)
        {
            #region Eg. “AC:pin_clk:tset5” will pick the period of tset5 of pin_clk
            if (_regex.IsMatch(newValue))
            {
                foreach (TSet timeSetData in timeSetSheet.Rows)
                {
                    if (!timeSetData.Name.Equals(newValue, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    foreach (TimingRow timeSetRow in timeSetData.TimingRows)
                    {
                        if (timeSetRow.PinGrpName.Equals(pinName, StringComparison.OrdinalIgnoreCase))
                        {
                            return timeSetRow.PinGrpClockPeriod;
                        }
                    }
                }
                return "";
            }
            #endregion

            double freq = double.Parse(HardIpPattern.GetFreq(newValue));
            return freq > 0 ? (1.0 / freq).ToString(CultureInfo.InvariantCulture) : "0";
        }

        internal void SetNewTimeSetValue(TimingRow timeSetRow, string newValue, bool isDiffPin, string cyclePeriod, string dataSrc, bool isPositivePin = true)
        {
            double.TryParse(newValue, out double value);
            bool isZero = false;
            if (Math.Abs(value) > 0 && value <= 1 / 500e6) // >500Mhz
            {
                SetClock2XTimeSetValue(timeSetRow, ref newValue, value, isDiffPin, isPositivePin);
            }
            else
            {
                SetClock1XTimeSetValue(timeSetRow, newValue, value, isDiffPin, cyclePeriod, isPositivePin, ref isZero);
            }
            if (!string.IsNullOrEmpty(dataSrc))
            {
                timeSetRow.DataSrc = dataSrc;
            }
            else
            {
                timeSetRow.DataSrc = isZero ? "ALLLO" : isDiffPin && !isPositivePin ? "ALLLO" : "ALLHI";
            }

            timeSetRow.CompareMode = "Off";
        }

        private void SetClock2XTimeSetValue(TimingRow timeSetRow, ref string newValue, double value, bool isDiffPin, bool isPositivePin)
        {
            newValue = (2 * value).ToString(CultureInfo.InvariantCulture);
            timeSetRow.PinGrpClockPeriod = newValue;
            timeSetRow.PinGrpSetup = Clock2X;
            timeSetRow.DataFmt = isDiffPin && !isPositivePin ? "RH-2X" : "RL-2X";
            timeSetRow.DriveOn = string.IsNullOrEmpty(newValue) ? newValue : "=" + newValue + "*1/4";
            timeSetRow.DriveData = string.IsNullOrEmpty(newValue) ? newValue : "=" + newValue + "*2/4";
            timeSetRow.DriveReturn = string.IsNullOrEmpty(newValue) ? newValue : "=" + newValue + "*3/4";
            timeSetRow.DriveOff = string.IsNullOrEmpty(newValue) ? newValue : "=" + newValue;
        }

        private void SetClock1XTimeSetValue(TimingRow timeSetRow, string newValue, double value, bool isDiffPin, string cyclePeriod, bool isPositivePin, ref bool isZero)
        {
            if (double.TryParse(cyclePeriod, out double valuePeriod))
            {
                if (!string.IsNullOrEmpty(timeSetRow.DriveOn) && double.TryParse(timeSetRow.DriveOn, out double valueOn))
                {
                    timeSetRow.DriveOn = "=" + newValue + "*" + (valueOn / valuePeriod);
                }

                if (!string.IsNullOrEmpty(timeSetRow.DriveData) &&
                    double.TryParse(timeSetRow.DriveData, out double valueData))
                {
                    timeSetRow.DriveData = "=" + newValue + "*" + (valueData / valuePeriod);
                }
                else
                {
                    timeSetRow.DriveData = string.IsNullOrEmpty(newValue) ? newValue : "=" + newValue + "*1/2";
                }

                if (!string.IsNullOrEmpty(timeSetRow.DriveReturn) &&
                    double.TryParse(timeSetRow.DriveReturn, out double valueReturn))
                {
                    timeSetRow.DriveReturn = "=" + newValue + "*" + (valueReturn / valuePeriod);
                }
                else
                {
                    timeSetRow.DriveReturn = newValue;
                }

                if (!string.IsNullOrEmpty(timeSetRow.DriveOff) &&
                    double.TryParse(timeSetRow.DriveOff, out double valueOff))
                {
                    timeSetRow.DriveOff = "=" + newValue + "*" + (valueOff / valuePeriod);
                }
            }
            else
            {
                timeSetRow.DriveData = string.IsNullOrEmpty(newValue) ? newValue : "=" + newValue + "*1/2";
                timeSetRow.DriveReturn = newValue;
            }
            isZero = !(Math.Abs(value) > 0);
            timeSetRow.PinGrpClockPeriod = isZero ? cyclePeriod : newValue;
            timeSetRow.PinGrpSetup = Clock;
            timeSetRow.DataFmt = isZero ? "RL" : isDiffPin && !isPositivePin ? "RH" : "RL";
        }

        private bool IsPositivePin(string diffPin, string pin)
        {
            DifferentialService.DiffPinPosAndNeg(diffPin, out string positivePin, out string _, out string _);
            if (positivePin.Equals(pin, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        internal void UpdateTimeSetSheetWithMcg(TimeSetBasicSheet timeSet, HardIpPattern pattern)
        {
            string timeSetName = timeSet.Name;

            if (string.IsNullOrEmpty(pattern.TimeSetUsed.McgSetting))
            {
                return;
            }

            #region check whether the pin exists in timeset sheet

            if (!ForceConditionChecker.CheckAcSettingPin(pattern, out string _))
            {
                return;
            }

            #endregion

            string timeSetUsedStr = _regex2.Replace(pattern.TimeSetUsed.McgSetting, "&");
            string timeUsedStr = ConvertUnits(timeSetUsedStr);

            #region if the MCG timeset was created, use it and will not create a new one
            string name = timeSetName.ToLower();
            if (_duplicateTimeSets.TryGetValue(name, out Dictionary<string, string> set))
            {
                string createdMcg = set.FirstOrDefault(s => s.Key.Equals(timeUsedStr, StringComparison.OrdinalIgnoreCase)).Value;
                if (createdMcg != null)
                {
                    pattern.TimeSetUsed.McgSetting = pattern.TimeSetUsed.McgSetting;
                    pattern.TimeSetUsed.TimeSet = timeSetName;
                    pattern.TimeSetUsed.TimeSetMcg = createdMcg;
                    return;
                }
            }
            #endregion

            #region if the MCG timeset was not created, create a new one: oldtimesetName + "_MCG_" + duplicate index
            TimeSetBasicSheet timeSetMcg = timeSet.Copy();
            string modifiedTimeSet = timeSetMcg.Name.Length > 23 ? timeSetMcg.Name.Substring(0, 23) : timeSetMcg.Name;
            string modifiedTimeSetName = modifiedTimeSet.ToLower();
            if (!_duplicateTimeSets.TryGetValue(modifiedTimeSetName, out Dictionary<string, string> dic))
            {
                var dictionary = new Dictionary<string, string>();
                timeSetMcg.Name = modifiedTimeSet + Mcg;
                dictionary.Add(timeUsedStr, timeSetMcg.Name);
                _duplicateTimeSets.Add(modifiedTimeSetName, dictionary);
            }
            else
            {
                timeSetMcg.Name = modifiedTimeSet + Mcg + "_" + _duplicateTimeSets[modifiedTimeSetName].Count;
                if (!dic.TryGetValue(timeUsedStr, out string value))
                {
                    dic.Add(timeUsedStr, timeSetMcg.Name);
                }
                else
                {

                    pattern.TimeSetUsed.McgSetting = pattern.TimeSetUsed.McgSetting;
                    pattern.TimeSetUsed.TimeSet = timeSetName;
                    pattern.TimeSetUsed.TimeSetMcg = value;
                    return;
                }
            }
            #endregion

            #region change pin clock period to new value specified in testplan
            if (_timeSetSheets.FirstOrDefault(s => s.Name.Equals(timeSetMcg.Name, StringComparison.OrdinalIgnoreCase)) == null)
            {
                foreach (string changedPin in timeSetUsedStr.Trim(';').Split(';'))
                {
                    string pinName = changedPin.Split(':')[1];
                    bool isDiffPin = pinName.Contains("&");
                    string value = changedPin.Split(':')[2];
                    string dataSrc = "";
                    if (value.Contains("@"))
                    {
                        string[] arr = value.Split('@');
                        dataSrc = arr[1].ToUpper();
                        value = arr[0];
                    }
                    string newValue = GetPinClockPeriod(pinName, value, timeSetMcg);
                    foreach (TSet timeSetData in timeSetMcg.Rows)
                    {
                        foreach (TimingRow timeSetRow in timeSetData.TimingRows)
                        {
                            if (!isDiffPin && timeSetRow.PinGrpName.Equals(pinName, StringComparison.OrdinalIgnoreCase))
                            {
                                SetNewTimeSetValue(timeSetRow, newValue, false, timeSetData.CyclePeriod, dataSrc);
                                continue;
                            }
                            if (pinName.Split('&').ToList().Exists(s => s.Equals(timeSetRow.PinGrpName, StringComparison.OrdinalIgnoreCase)))
                            {
                                bool isPositive = IsPositivePin(pinName.Replace("&", "::"), timeSetRow.PinGrpName);
                                SetNewTimeSetValue(timeSetRow, newValue, true, timeSetData.CyclePeriod, dataSrc, isPositive);
                            }
                        }
                    }
                }
                _timeSetSheets.Add(timeSetMcg);
                if (AcTSetCategoryMapSingleton.Instance().TryContains(timeSet.Name, out TimeSetBlock2Category originalTimeSet))
                {
                    AcTSetCategoryMapSingleton.Instance().SetRow(timeSetMcg.Name, originalTimeSet.Block, originalTimeSet.Category);
                }
            }
            #endregion

            pattern.TimeSetUsed.McgSetting = pattern.TimeSetUsed.McgSetting;
            pattern.TimeSetUsed.TimeSet = timeSetName;
            pattern.TimeSetUsed.TimeSetMcg = timeSetMcg.Name;
        }
    }
}
