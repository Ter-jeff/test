using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.Static;

using CommonLib.Extension;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class MeasPinCurrentRangeCalculator
    {
        private static readonly Regex _regex = new Regex("@\"(?<Value1>[+-]?\\d+([.])?(\\d+)?)[\\+\\-\\*\\\\]?(?<Value2>[+-]?(\\d+)?[.]?(\\d+)?)(?<Unit>\\w+)*\"", RegexOptions.Compiled);

        private readonly MeasPin _pin;

        private List<ForcePin> _postCachedForcePins;
        private List<double> _postCachedForceValues;

        public MeasPinCurrentRangeCalculator(MeasPin pin)
        {
            _pin = pin;
        }

        public List<CurrentRange> GetCurrentRangeListByVoltage(List<MeasLimit> measLimits)
        {
            EnsurePostForceCache();

            List<ForcePin> forcePinlist = _postCachedForcePins;
            List<double> forceValueList = _postCachedForceValues;

            List<CurrentRange> currentList = new List<CurrentRange>();
            if (_pin.MeasType.Equals("measi", StringComparison.OrdinalIgnoreCase) ||
                _pin.MeasType.Equals("measidiff", StringComparison.OrdinalIgnoreCase) ||
                Regex.IsMatch(_pin.MeasType, "MeasR[1|2]", RegexOptions.IgnoreCase))
            {


                for (int i = 0; i < _pin.MeasLimitsH.Count; i++)
                {
                    var highLimitList = new List<string>();
                    var lowLimitList = new List<string>();

                    GetLimitList(measLimits[i], ref lowLimitList, ref highLimitList);

                    double hiRange = ParseLimitList(highLimitList);
                    double loRange = ParseLimitList(lowLimitList);

                    if (!Regex.IsMatch(_pin.MeasType, @"^MeasR\d"))
                    {
                        CurrentRange range = new CurrentRange
                        {
                            JobName = _pin.MeasLimitsL[i].JobName,
                            Value = hiRange == 0 && loRange == 0 ? "999.999" : (hiRange > loRange ? hiRange : loRange).ToString("G15")
                        };
                        currentList.Add(range);
                    }
                    else
                    {
                        if (!(hiRange > 0 || loRange > 0))
                        {
                            continue;
                        }

                        if (_pin.ForceConditions.Count == 0)
                        {
                            return null;
                        }

                        if (forceValueList.Count == 0 || forceValueList.Max() == 0)
                        {
                            return null;
                        }

                        #region if forceType == "I", currentRange = forceI value

                        if (forcePinlist.All(x => x.ForceType.Equals("I", StringComparison.OrdinalIgnoreCase)))
                        {
                            CurrentRange range = new CurrentRange
                            {
                                JobName = _pin.MeasLimitsL[i].JobName,
                                Value = Math.Abs(forceValueList.Max()).ToString("G15")
                            };
                            currentList.Add(range);
                        }
                        #endregion

                        #region if forceType == "V", currentRange = (forceV value)/lowest(limit range)
                        else if (forcePinlist.All(x => x.ForceType.Equals("V", StringComparison.OrdinalIgnoreCase)))
                        {
                            double value;
                            if (hiRange > 0 && loRange > 0)
                            {
                                value = Math.Abs(hiRange > loRange ? loRange : hiRange); // get max i range with max voltage and min R
                            }
                            else
                            {
                                value = Math.Abs(hiRange > loRange ? hiRange : loRange); //get value that is not zero
                            }

                            CurrentRange range = new CurrentRange
                            {
                                JobName = _pin.MeasLimitsL[i].JobName,
                                Value = Math.Abs(Math.Round(forceValueList.Max() / value, 7)).ToString("G15")
                            };
                            currentList.Add(range);
                        }
                        #endregion
                    }
                }
                return currentList;
            }
            return null;
        }

        private void EnsurePostForceCache()
        {
            if (_postCachedForcePins == null)
            {
                _postCachedForcePins = GestForcePinlist();
            }

            if (_postCachedForceValues == null)
            {
                _postCachedForceValues = GetForceValueList(_postCachedForcePins);
            }
        }

        public List<CurrentRange> GetCurrentRangeList()
        {
            EnsurePostForceCache();

            List<ForcePin> forcePinlist = _postCachedForcePins;
            List<double> forceValueList = _postCachedForceValues;

            List<CurrentRange> currentList = new List<CurrentRange>();
            if (_pin.MeasType.Equals("measi", StringComparison.OrdinalIgnoreCase) ||
                _pin.MeasType.Equals("measidiff", StringComparison.OrdinalIgnoreCase) ||
                Regex.IsMatch(_pin.MeasType, "MeasR[1|2]", RegexOptions.IgnoreCase))
            {


                for (int i = 0; i < _pin.MeasLimitsH.Count; i++)
                {
                    var highLimitList = new List<string>();
                    var lowLimitList = new List<string>();

                    GetLimitList(_pin.MeasLimitsH[i], ref lowLimitList, ref highLimitList);
                    GetLimitList(_pin.MeasLimitsL[i], ref lowLimitList, ref highLimitList);
                    GetLimitList(_pin.MeasLimitsN[i], ref lowLimitList, ref highLimitList);

                    double hiRange = ParseLimitList(highLimitList);
                    double loRange = ParseLimitList(lowLimitList);

                    if (!Regex.IsMatch(_pin.MeasType, @"^MeasR\d"))
                    {
                        CurrentRange range = new CurrentRange
                        {
                            JobName = _pin.MeasLimitsL[i].JobName,
                            Value = hiRange == 0 && loRange == 0 ? "999.999" : (hiRange > loRange ? hiRange : loRange).ToString("G15")
                        };
                        currentList.Add(range);
                    }
                    else
                    {
                        if (!(hiRange > 0 || loRange > 0))
                        {
                            continue;
                        }

                        if (_pin.ForceConditions.Count == 0)
                        {
                            return null;
                        }

                        if (forceValueList.Count == 0 || forceValueList.Max() == 0)
                        {
                            return null;
                        }

                        #region if forceType == "I", currentRange = forceI value

                        if (forcePinlist.All(x => x.ForceType.Equals("I", StringComparison.OrdinalIgnoreCase)))
                        {
                            CurrentRange range = new CurrentRange
                            {
                                JobName = _pin.MeasLimitsL[i].JobName,
                                Value = Math.Abs(forceValueList.Max()).ToString("G15")
                            };
                            currentList.Add(range);
                        }
                        #endregion

                        #region if forceType == "V", currentRange = (forceV value)/lowest(limit range)
                        else if (forcePinlist.All(x => x.ForceType.Equals("V", StringComparison.OrdinalIgnoreCase)))
                        {
                            double value;
                            if (hiRange > 0 && loRange > 0)
                            {
                                value = Math.Abs(hiRange > loRange ? loRange : hiRange); // get max i range with max voltage and min R
                            }
                            else
                            {
                                value = Math.Abs(hiRange > loRange ? hiRange : loRange); //get value that is not zero
                            }

                            CurrentRange range = new CurrentRange
                            {
                                JobName = _pin.MeasLimitsL[i].JobName,
                                Value = Math.Abs(Math.Round(forceValueList.Max() / value, 7)).ToString("G15")
                            };
                            currentList.Add(range);
                        }
                        #endregion
                    }
                }
                return currentList;
            }
            return null;
        }

        private List<ForcePin> GestForcePinlist()
        {
            List<ForcePin> forcePinlist = new List<ForcePin>();
            foreach (ForceCondition forceCondtion in _pin.ForceConditions)
            {
                foreach (ForcePin forcePin in forceCondtion.ForcePins)
                {
                    if (forcePin.PinName.Equals(_pin.PinName, StringComparison.OrdinalIgnoreCase))
                    {
                        forcePinlist.Add(forcePin);
                    }
                    else
                    {
                        //DecompGroups for force condition and Misc info match
                        List<string> foccenameList = new List<string>();
                        foccenameList.AddRange(TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(forcePin.PinName));

                        List<ForcePin> newforcePinList = new List<ForcePin>();
                        foreach (string a in foccenameList)
                        {
                            ForcePin tempforcePin = forcePin.Copy();
                            tempforcePin.PinName = a;
                            newforcePinList.Add(tempforcePin);
                        }

                        ForcePin newforcePin = newforcePinList.Find(s => s.PinName.Equals(_pin.PinName, StringComparison.OrdinalIgnoreCase));
                        if (newforcePin != null)
                        {
                            forcePinlist.Add(newforcePin);
                        }
                    }
                }
            }
            return forcePinlist;
        }

        private static List<double> GetForceValueList(List<ForcePin> forcePinlist)
        {
            List<double> forceValueList = new List<double>();
            foreach (ForcePin forcePin in forcePinlist)
            {
                foreach (string value in forcePin.ForceValue.Split('&'))
                {
                    if (double.TryParse(value, out double forceValue))
                    {
                        forceValueList.Add(forceValue);
                    }
                    else
                    {
                        return forceValueList;
                    }
                }
            }
            return forceValueList;
        }

        private string ConvertToNum(string limit)
        {
            Match match = _regex.Match(limit);
            string value1 = match.Groups["Value1"].ToString();
            string value2 = match.Groups["Value2"].ToString();
            string unit = match.Groups["Unit"].ToString();

            if (!double.TryParse(value1, out double tmpValue1))
            {
                return limit;
            }

            if (!double.TryParse(value2, out double tmpValue2))
            {
                return limit;
            }

            double value = tmpValue1 * tmpValue2;

            return value + unit;
        }

        private void GetLimitList(MeasLimit limit, ref List<string> loValueList, ref List<string> hiValueList)
        {
            string lowLimit = DataConvertor.ConvertUnits(ConvertToNum(limit.LoLimit));
            if (Regex.IsMatch(lowLimit, @"^(\d|\.|-)$") || lowLimit.Contains("E-") ||
                Regex.IsMatch(lowLimit, @"^(\d|\.|-)+$") || lowLimit.Equals(""))
            {
                loValueList.Add(lowLimit);
            }

            string highLimit = DataConvertor.ConvertUnits(ConvertToNum(limit.HiLimit));
            if (Regex.IsMatch(highLimit, @"^(\d|\.|-)$") || highLimit.Contains("E-") ||
                Regex.IsMatch(highLimit, @"^(\d|\.|-)+$") || highLimit.Equals(""))
            {
                hiValueList.Add(highLimit);
            }
        }

        private double ParseLimitList(List<string> limitList)
        {
            double range = 0;
            if (limitList == null || limitList.Count == 0)
            {
                return range;
            }

            bool anyNonEmpty = false;
            double minAbs = double.MaxValue;
            double maxAbs = 0;

            for (int i = 0; i < limitList.Count; i++)
            {
                string s = limitList[i];
                if (string.IsNullOrEmpty(s))
                {
                    continue;
                }

                anyNonEmpty = true;

                double val = Convert.ToDouble(s);
                double absVal = Math.Abs(val);
                if (absVal < minAbs)
                {
                    minAbs = absVal;
                }

                if (absVal > maxAbs)
                {
                    maxAbs = absVal;
                }
            }

            if (anyNonEmpty)
            {
                range = _pin.MeasType.ContainsIgnoreCase("R") ? minAbs : maxAbs;
            }

            return range;
        }

        public string GetCurrentRangeByVoltage(string voltage)
        {
            List<CurrentRange> currentRangeList;
            switch (voltage)
            {
                case "HV":
                    currentRangeList = _pin.CurrentRangeListH;
                    break;
                case "LV":
                    currentRangeList = _pin.CurrentRangeListL;
                    break;
                case "NV":
                    currentRangeList = _pin.CurrentRangeListN;
                    break;
                default:
                    currentRangeList = _pin.CurrentRangeList;
                    break;
            }
            if (currentRangeList == null || currentRangeList.Count == 0)
            {
                return "";
            }

            if (_pin.PinName.Contains('='))
            {
                string stage = _pin.PinName.Split('=').First();
                var joblist = new List<string> { stage + "1:", stage + "2:" };
                if (!currentRangeList.Any())
                {
                    return string.Join(";", joblist);
                }
                return string.Join(";",
                    currentRangeList.Where(x => x.JobName.Contains(stage)).Select(x => x.JobName + ":" + x.Value));
            }

            if (currentRangeList.Select(x => x.Value).Distinct().Count() == 1 && !_pin.PinName.Contains("="))
            {
                return currentRangeList[0].Value;
            }

            return string.Join(";", currentRangeList.Select(x => x.JobName + ":" + x.Value));
        }

        public string GetCurrentRangeByVoltageCp(string voltage)
        {
            List<CurrentRange> currentRangeList;
            switch (voltage)
            {
                case "HV":
                    currentRangeList = _pin.CurrentRangeListH;
                    break;
                case "LV":
                    currentRangeList = _pin.CurrentRangeListL;
                    break;
                case "NV":
                    currentRangeList = _pin.CurrentRangeListN;
                    break;
                default:
                    currentRangeList = _pin.CurrentRangeList;
                    break;
            }
            var joblist = new List<string> { "CP1:", "CP2:" };
            if (!currentRangeList.Any())
            {
                return string.Join(";", joblist);
            }
            return string.Join(";",
                currentRangeList.Where(x => x.JobName.Contains("CP")).Select(x => x.JobName + ":" + x.Value));
        }

        public string GetCurrentRange()
        {
            if (_pin.CurrentRangeList == null || _pin.CurrentRangeList.Count == 0)
            {
                return "";
            }

            if (_pin.CurrentRangeList.Select(x => x.Value).Distinct().Count() == 1)
            {
                return _pin.CurrentRangeList[0].Value;
            }

            return string.Join(";", _pin.CurrentRangeList.Select(x => x.JobName + ":" + x.Value));
        }
    }
}
