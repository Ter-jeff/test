using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.PowerMerge;
using TestPlanLib.Utility;

namespace Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility
{
    public class DataConvertor
    {
        private static readonly Regex _reg5 = new Regex(@"[\w|.]+", RegexOptions.Compiled);
        private static readonly Regex _reg6 = new Regex(@"[-|\w|.]+", RegexOptions.Compiled);
        private static readonly Regex _reg7 = new Regex(@"(?<spec>\w+)\s*(\((?<name>\w+)\))*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex _regex = new Regex(@"^(\d|\.|-)+$", RegexOptions.Compiled);
        private static readonly Regex _regex21 = new Regex(@"^(\d|\.|-)+(\w)*$", RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(@"^(\d|\.|-)+(\w|%)*$", RegexOptions.Compiled);
        private static readonly Regex _regex3 = new Regex(@"(?<num>((\d|\.|-)+))[^\*]*", RegexOptions.Compiled);
        private static readonly Regex _regex4 = new Regex(@"\++", RegexOptions.Compiled);
        private static readonly Regex _regex6 = new Regex(@"[\w|.|%|""]+", RegexOptions.Compiled);
        private static readonly Regex _regex7 = new Regex(@".*\w+[\+|\-|\*|\/]\w+", RegexOptions.Compiled);
        private static readonly Regex _regex1 = new Regex(@"^(\d|\.|-)+(\w|%|"")*$", RegexOptions.Compiled);
        private static readonly Regex _regex9 = new Regex(@"^\w*$", RegexOptions.Compiled);
        private static readonly Regex _regex11 = new Regex("^%.*", RegexOptions.Compiled);
        private static readonly Regex _regex12 = new Regex("^m.*", RegexOptions.Compiled);
        private static readonly Regex _regex13 = new Regex("^u.*", RegexOptions.Compiled);
        private static readonly Regex _regex14 = new Regex("^n.*", RegexOptions.Compiled);
        private static readonly Regex _regex15 = new Regex("^p.*", RegexOptions.Compiled);
        private static readonly Regex _regex16 = new Regex("^k.*", RegexOptions.Compiled);
        private static readonly Regex _regex17 = new Regex("^M.*", RegexOptions.Compiled);
        private static readonly Regex _regex18 = new Regex("^G.*", RegexOptions.Compiled);
        private static readonly Regex _regex19 = new Regex("^T.*", RegexOptions.Compiled);
        private static readonly Regex _regex20 = new Regex("^\".*", RegexOptions.Compiled);
        private static readonly Regex _regex22 = new Regex("^No.*", RegexOptions.Compiled);

        public const string Var = "_VAR";

        public static string ConvertUnits(string limitStr, bool nonScience = false)
        {
            if (limitStr.Contains("10^"))
            {
                limitStr = limitStr.Replace("*10^", "E");
            }

            if (limitStr == "" || limitStr.Contains("E") || _regex.IsMatch(limitStr))
            {
                return limitStr;
            }

            if (_regex21.IsMatch(limitStr))
            {
                string limitNum = _regex3.Match(limitStr).Groups["num"].ToString();
                if (limitNum == "0")
                {
                    return limitNum;
                }

                string limitUnit = limitStr.Replace(limitNum, "").Trim();
                double rate = 1;
                if (_regex12.IsMatch(limitUnit))
                {
                    rate = 1 / (double)1000;
                }
                else if (_regex13.IsMatch(limitUnit))
                {
                    rate = 1 / (double)1000000;
                }
                else if (_regex14.IsMatch(limitUnit))
                {
                    rate = 1 / (double)1000000000;
                }
                else if (_regex15.IsMatch(limitUnit))
                {
                    rate = 1 / (double)1000000000000;
                }
                else if (_regex16.IsMatch(limitUnit.ToLower()))
                {
                    rate = 1000;
                }
                else if (_regex17.IsMatch(limitUnit))
                {
                    rate = 1000000;
                }
                else if (_regex18.IsMatch(limitUnit))
                {
                    rate = 1000000000;
                }
                else if (_regex19.IsMatch(limitUnit))
                {
                    rate = 1000000000000;
                }

                if (double.TryParse(limitNum, out double value))
                {
                    return nonScience ? (value * rate).ToString("0." + new string('#', 339), CultureInfo.InvariantCulture) : (value * rate).ToString("G15", CultureInfo.InvariantCulture);
                }
            }
            return limitStr;
        }

        internal static string ConvertUnits(string limitStr, out string limitUnit, out string limitScale)
        {
            limitUnit = "";
            limitScale = "";

            if (limitStr == "" || limitStr.Contains("E") || _regex.IsMatch(limitStr))
            {
                return limitStr;
            }

            if (_regex1.IsMatch(limitStr))
            {
                string limitNum = _regex3.Match(limitStr).Groups["num"].ToString();
                double rate;

                limitUnit = limitStr.Replace(limitNum, "").Trim();
                (rate, limitUnit, limitScale) = GetRateUnitScale(limitUnit);
                limitUnit = limitUnit.ToUpper();
                if (limitUnit == "HZ")
                {
                    limitUnit = "Hz";
                }
                else if (limitUnit == "OHM")
                {
                    limitUnit = "Ohm";
                }
                else if (limitUnit == "OHMS")
                {
                    limitUnit = "Ohms";
                }
                else if (limitUnit == "CODE")
                {
                    limitUnit = "code";
                }

                if (double.TryParse(limitNum, out double value))
                {
                    return (value * rate).ToString("G15", CultureInfo.InvariantCulture);
                }
            }
            else if (limitStr.StartsWith("NA", StringComparison.CurrentCultureIgnoreCase))
            {
                bool isNa = limitStr.Substring(0, 2) == "NA";
                if (isNa)
                {
                    limitUnit = limitStr.Substring(2, limitStr.Length - 2);
                    limitStr = "";
                    (_, limitUnit, limitScale) = GetRateUnitScale(limitUnit);
                    limitUnit = limitUnit.ToUpper();
                    if (limitUnit == "HZ")
                    {
                        limitUnit = "Hz";
                    }
                    else if (limitUnit == "OHM")
                    {
                        limitUnit = "Ohm";
                    }
                    else if (limitUnit == "OHMS")
                    {
                        limitUnit = "Ohms";
                    }

                    if (!string.IsNullOrEmpty(limitScale))
                    {
                        limitStr = "";
                    }
                }
            }
            else if (_regex9.IsMatch(limitStr))
            {
                (_, limitUnit, limitScale) = GetRateUnitScale(limitStr);
                return "";
            }
            return limitStr;
        }

        internal static (double, string, string) GetRateUnitScale(string limitUnit)
        {
            double rate = 1;
            string limitScale = "";
            if (_regex11.IsMatch(limitUnit))
            {
                rate = 1 / (double)100;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScalePercent;
            }
            else if (_regex12.IsMatch(limitUnit))
            {
                rate = 1 / (double)1000;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScaleMilli;
            }
            else if (_regex13.IsMatch(limitUnit))
            {
                rate = 1 / (double)1000000;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScaleMicro;
            }
            else if (_regex14.IsMatch(limitUnit))
            {
                rate = 1 / (double)1000000000;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScaleNano;
            }
            else if (_regex15.IsMatch(limitUnit))
            {
                rate = 1 / (double)1000000000000;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScalePico;
            }
            else if (_regex16.IsMatch(limitUnit.ToLower()))
            {
                rate = 1000;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScaleKilo;
            }
            else if (_regex17.IsMatch(limitUnit))
            {
                rate = 1000000;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScaleMega;
            }
            else if (_regex18.IsMatch(limitUnit))
            {
                rate = 1000000000;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScaleGiga;
            }
            else if (_regex19.IsMatch(limitUnit))
            {
                rate = 1000000000000;
                limitUnit = RemoveScale(limitUnit);
                limitScale = CommonConst.ScaleTera;
            }
            else if (_regex20.IsMatch(limitUnit))
            {
                limitUnit = limitUnit.Trim('\"');
            }
            else if (_regex22.IsMatch(limitUnit))
            {
                limitUnit = limitUnit.Substring(2, limitUnit.Length - 2);
                limitScale = CommonConst.ScaleNo;
            }
            return (rate, limitUnit, limitScale);
        }

        private static string RemoveScale(string limitUnit)
        {
            return limitUnit.Substring(1, limitUnit.Length - 1);
        }

        public static string ConvertUseLimit(string limitStr, out string limitUnit, out string limitScale, string chiplet = "")
        {
            if (limitStr.StartsWith("Binning", StringComparison.CurrentCultureIgnoreCase))
            {
                limitUnit = "";
                limitScale = "";
                return "";
            }

            if (ConvertNumber(limitStr, out long value))
            {
                limitUnit = "";
                limitScale = "";
                return value.ToString();
            }

            return ConvertUseLimitToGlbSpec(limitStr, out limitUnit, out limitScale, chiplet);
        }

        public static string ConvertUseLimitFw(string limitStr, out string limitUnit, out string limitScale)
        {
            limitUnit = "";
            limitScale = "";

            if (limitStr == "" || limitStr.Contains("E") || _regex.IsMatch(limitStr))
            {
                return limitStr;
            }

            if (_regex2.IsMatch(limitStr))
            {
                string limitNum = _regex3.Match(limitStr).Groups["num"].ToString();

                limitUnit = limitStr.Replace(limitNum, "").Trim();
                limitStr = limitStr.Replace(limitUnit, "");
            }
            else if (limitStr.StartsWith("NA", StringComparison.CurrentCultureIgnoreCase))
            {
                bool isNa = limitStr.Substring(0, 2) == "NA";
                if (isNa)
                {
                    limitUnit = limitStr.Substring(2, limitStr.Length - 2);
                    limitStr = "";
                }
            }
            return limitStr;
        }

        private static string ConvertUseLimitToGlbSpec(string limitStr, out string limitUnit, out string limitScale, string chiplet = "")
        {
            limitUnit = "";
            limitScale = "";

            MatchCollection matches = _regex6.Matches(limitStr);
            string result = limitStr;
            List<string> list = new List<string>();
            foreach (Match m in matches)
            {
                if (m.Value.Trim().ToUpper().StartsWith("VDD", StringComparison.OrdinalIgnoreCase))
                {
                    if (!list.Contains(m.Value))
                    {
                        if (MultiTestSettingUtility.ExistChipletVddPin(m.Value, chiplet, TestPlanStatic.PowerInfoSheet))
                        {
                            result = result.Replace(m.Value, "_" + m.Value.ToUpper() + Var + "_" + chiplet);
                        }
                        else
                        {
                            result = result.Replace(m.Value, "_" + m.Value.ToUpper() + Var);
                        }

                        list.Add(m.Value);
                    }
                }
                else
                {
                    if (!list.Contains(m.Value))
                    {
                        result = result.Replace(m.Value, ConvertUnits(m.Value, out limitUnit, out limitScale));
                        list.Add(m.Value);
                    }
                }
            }
            if (result.Contains(Var) || _regex7.IsMatch(result))
            {
                result = "=" + result;
            }

            return result;
        }

        private static bool ConvertNumber(string text, out long value)
        {
            value = 0;
            if (text.Length <= 2)
            {
                return false;
            }

            string prefix = text.Substring(0, 2).ToLower();
            string number = text.Remove(0, 2);
            try
            {
                switch (prefix)
                {
                    case "0b":
                        value = Convert.ToInt64(number, 2);
                        return true;
                    case "0x":
                        value = Convert.ToInt64(number, 16);
                        return true;
                    case "0d":
                        value = Convert.ToInt64(number);
                        return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return false;
        }

        public static string ConvertValueSpec(string value, string voltage = "")
        {
            if (value.ContainsIgnoreCase("reg_assign"))
            {
                return value;
            }

            var list = new List<string>();
            try
            {

                if (value != "")
                {
                    foreach (string subvalue in value.Split(';'))
                    {
                        bool isAlg = subvalue.ContainsIgnoreCase("Alg:");
                        if (!isAlg)
                        {
                            List<string> items = subvalue.Replace(" ", "").Split(':').ToList();
                            items[2] = _reg7.Replace(items[2], delegate (Match m)
                            {
                                if (!string.IsNullOrEmpty(m.Groups["name"].Value) || !TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(m.Groups["spec"].Value.ToUpper()))
                                {
                                    return m.Value;
                                }

                                if (m.Value == m.Groups["spec"].Value)
                                {
                                    return "_" + m.Groups["spec"].Value.ToUpper() + Var;
                                }

                                return "_" + m.Groups["spec"].Value.ToUpper() + Var + m.Value[m.Value.Length - 1];
                            });
                            list.Add(string.Join(":", items));
                        }
                        else
                        {
                            string newValue = subvalue;
                            if (!string.IsNullOrEmpty(voltage))
                            {
                                Match getAlg = Regex.Match(newValue, $"^(?<vol>({voltage})@)?Alg:", RegexOptions.IgnoreCase);
                                if (getAlg.Success)
                                {
                                    string algVol = getAlg.Groups["vol"].Value;
                                    if (!string.IsNullOrEmpty(algVol))
                                    {
                                        newValue = Regex.Replace(newValue, $"^{algVol}", "");
                                    }
                                    list.Add(newValue);
                                }
                            }
                            else
                            {
                                if (newValue.StartsWith("Alg:", StringComparison.OrdinalIgnoreCase))
                                {
                                    list.Add(newValue);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                //pass;
            }
            return string.Join(";", list);
        }

        public static string ConvertForceValueToGlbSpec(ForcePin forcePin, string chiplet = "", bool nonScience = false)
        {
            string forceValue = forcePin.ForceValue;
            return _reg5.Replace(forceValue, delegate (Match m)
            {
                if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(m.Value.ToUpper()))
                {
                    return ConvertUnits(m.Value, nonScience);
                }

                return "_" + m.Value.ToUpper() + Var;
            });
        }

        public static string ConvertValueWithGlbSpec(string value, bool nonScience = false)
        {
            string forceValue = value;
            return _reg6.Replace(forceValue, delegate (Match m)
            {
                if (!TestProgram.IgxlWorkBk.PinMapPair.Value.IsPinExist(m.Value.ToUpper()))
                {
                    return ConvertUnits(m.Value, nonScience);
                }

                return "_" + m.Value.ToUpper() + Var;
            });
        }

        /// <summary>
        /// Convert bump name and ball name to Net Name for MeasStr in patInfo
        /// </summary>
        /// <param name="measStr"></param>
        /// <returns></returns>
        public static string ConvertToNetName(string measStr, PowerMerge powerMerge)
        {
            string resultStr;
            if (!measStr.Contains("VDD", StringComparison.OrdinalIgnoreCase))
            {
                resultStr = measStr;
            }
            else
            {
                string newMeasStr = "";
                foreach (string seq in measStr.Split('+'))
                {
                    string newSeq = "";
                    if (seq != "")
                    {
                        var newSeqList = new List<string>();
                        foreach (string pinName in seq.Split(','))
                        {
                            string cpNetName = "";
                            string ftNetName = "";

                            if (powerMerge == null)
                            {
                                return measStr;
                            }

                            powerMerge.GetCpFtNetName(pinName, ref cpNetName, ref ftNetName);
                            if (cpNetName == "" && ftNetName == "")
                            {
                                newSeqList.Add(pinName);
                            }
                            else if (cpNetName == ftNetName)
                            {
                                newSeqList.Add(cpNetName);
                            }
                            else if (cpNetName != ftNetName)
                            {
                                if (cpNetName != "" && cpNetName != "N/A" && !newSeqList.Contains(cpNetName))
                                {
                                    newSeqList.Add("CP=" + cpNetName.Replace(",", ",CP="));
                                }
                                if (ftNetName != "" && ftNetName != "N/A" && !newSeqList.Contains(ftNetName))
                                {
                                    newSeqList.Add("FT=" + ftNetName.Replace(",", ",FT="));
                                }
                            }
                        }
                        newSeq = string.Join(",", newSeqList.Distinct());
                    }
                    newMeasStr += newSeq + "+";
                }
                resultStr = newMeasStr.Remove(newMeasStr.Length - 1, 1);
            }
            return resultStr;
        }

        public static string SortCpFtPin(string measStr)
        {
            bool needSort = false;
            List<string> seqList = measStr.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            seqList.ForEach(p =>
            {
                if (p.Split(',').ToList().Exists(k => k.Contains("CP=") || k.Contains("FT=")))
                {
                    needSort = true;
                }
            });

            if (!needSort)
            {
                return measStr;
            }

            var newCpSeqList = new List<string>();
            var newFtSeqList = new List<string>();
            List<string> measSeqList = measStr.Split('+').ToList();
            foreach (string seqPin in measSeqList)
            {
                var cpMeasList = new List<string>();
                var ftMeasList = new List<string>();
                if (string.IsNullOrWhiteSpace(seqPin))
                {
                    newCpSeqList.Add("");
                    newFtSeqList.Add("");
                    continue;
                }
                if (!seqPin.Contains("CP=") && !seqPin.Contains("FT="))
                {
                    newCpSeqList.Add(seqPin);
                    newFtSeqList.Add(seqPin);
                }
                else
                {
                    if (seqPin.Contains(":"))
                    {
                        foreach (string pin in seqPin.Split(','))
                        {
                            foreach (string sweepPin in pin.Split(':'))
                            {
                                if (sweepPin.Contains("CP="))
                                {
                                    cpMeasList.Add(sweepPin.Replace("CP=", ""));
                                }
                                else if (sweepPin.Contains("FT="))
                                {
                                    ftMeasList.Add(sweepPin.Replace("FT=", ""));
                                }
                                else
                                {
                                    cpMeasList.Add(sweepPin);
                                    ftMeasList.Add(sweepPin);
                                }
                            }
                        }
                        newCpSeqList.Add(string.Join(":", cpMeasList));
                        newFtSeqList.Add(string.Join(":", ftMeasList));
                    }
                    else
                    {
                        foreach (string pin in seqPin.Split(','))
                        {
                            if (pin.Contains("CP="))
                            {
                                cpMeasList.Add(pin.Replace("CP=", ""));
                            }
                            else if (pin.Contains("FT="))
                            {
                                ftMeasList.Add(pin.Replace("FT=", ""));
                            }
                            else
                            {
                                cpMeasList.Add(pin);
                                ftMeasList.Add(pin);
                            }
                        }
                        IEnumerable<string> tmpCpMeasList = cpMeasList.Distinct();
                        IEnumerable<string> tmpFtMeasList = ftMeasList.Distinct();

                        newCpSeqList.Add(string.Join(",", tmpCpMeasList));
                        newFtSeqList.Add(string.Join(",", tmpFtMeasList));
                    }
                }
            }
            return "CP=" + string.Join("+", newCpSeqList) + ";FT=" + string.Join("+", newFtSeqList);
        }

        /// <summary>
        /// Remove jobs if all jobs are enable
        /// </summary> If the job enable is "CP1,CP2,FT1,FT2,FT3", leave it as empty
        /// <param name="flowSheet"></param>
        public static void FilterFlowJobs(SubFlowSheet flowSheet)
        {
            foreach (FlowRow row in flowSheet.Rows)
            {
                if (!string.IsNullOrEmpty(row.Job))
                {
                    List<string> jobs = row.Job.ToUpper().Split(',').ToList();
                    if (jobs.All(LocalSpecs.AllJobsHardIp.Contains) && LocalSpecs.AllJobsHardIp.All(jobs.Contains))
                    {
                        row.Job = "";
                    }
                }
            }
        }

        /// <summary>
        /// If Measure pins are the same in different sequence, just keep one. e.g.  PinA,PinB+PinA,PinB+PinA,PinB ==> PinA,PinB
        /// </summary>
        /// <param name="measPins"></param>
        /// <returns></returns>
        public static string RemoveDummyPlusSign(string measPins)
        {
            List<string> pins = _regex4.Split(measPins).ToList();
            var mergedPin = pins.Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
            if (mergedPin.Count == 1)
            {
                return mergedPin.FirstOrDefault()?.Trim(',');
            }

            return measPins.Trim(',');
        }

        /// <summary>
        /// Convert differential pins to pin group. e.g. "Pin_Diff" ==》 "Pin_P,Pin_N"
        /// </summary>
        /// <param name="measPins"></param>
        /// <returns></returns>
        public static string ConvertDifferentialPinGroup(string measPins)
        {
            string result = measPins;
            if (!measPins.Contains("::") && TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(measPins))
            {
                result = "";
                List<string> pinList = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(measPins);
                for (int i = 0; i < pinList.Count; i++)
                {
                    result += pinList[i + 1] + "::" + pinList[i] + ",";
                    i++;
                }
                result = result.TrimEnd(',');
            }

            return result;
        }
    }
}
