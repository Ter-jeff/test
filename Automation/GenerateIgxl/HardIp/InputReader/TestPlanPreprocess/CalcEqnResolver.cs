using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using LogLib.Utility;

namespace Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    internal class CalcEqnResolver
    {
        private const string RegScientificNotation = @"\((?<value>\d*(?:\.\d*)?)[Ee](?<power>[+-]\d*)\)"; //(10E+8) or (10E-6)
        private const string RegPowerNotation = @"\((?<value>\d*)[\^](?<power>\d*)\)"; //(X^Y)

        private readonly TestPlanSheet _planSheet;
        private readonly Dictionary<string, List<string>> _allCusStr = new Dictionary<string, List<string>>();

        public CalcEqnResolver(TestPlanSheet planSheet)
        {
            _planSheet = planSheet;
        }

        internal string GetPatternCalcEqn(HardIpPattern pattern)
        {
            var calcEqnlst = new List<string>();
            var tmpMeasPins = new List<MeasPin>();
            for (int i = 0; i < pattern.MeasPins.Count; i++)
            {
                ProcessMeasPinCalcEqn(pattern, i, calcEqnlst, tmpMeasPins);
            }
            pattern.MeasPins = tmpMeasPins;
            string patternCalcEqn = GetCalcEqnForPattern(pattern.MiscInfo);
            if (!calcEqnlst.Contains(patternCalcEqn) && patternCalcEqn != "")
            {
                IEnumerable<MeasPin> calcEqnLimits = pattern.MeasPins.Where(x => x.MeasType == MeasType.MeasLimit || x.MeasType == MeasType.MeasCalc);
                if (calcEqnLimits.Any() && calcEqnLimits.FirstOrDefault().MeasType == MeasType.MeasLimit)
                {
                    var tmp = new List<string>(calcEqnlst);
                    calcEqnlst = JudgeCalcEqnByJob(patternCalcEqn);
                    calcEqnlst.AddRange(tmp);
                }
                else
                {
                    calcEqnlst.AddRange(JudgeCalcEqnByJob(patternCalcEqn));
                }
            }
            string regCalcReplace = @"Calc_Eqn:[""](?<content>.*)[""]";
            // force overwrite if misc info use calc_eqn to replace tradition syntax
            if (pattern.MiscInfo.ContainsIgnoreCase("calc_eqn"))
            {
                string content = Regex.Match(pattern.MiscInfo, regCalcReplace, RegexOptions.IgnoreCase).Groups["content"].Value;
                pattern.MiscInfo = Regex.Replace(pattern.MiscInfo, regCalcReplace, "", RegexOptions.IgnoreCase);
                return content;
            }
            return string.Join(";", calcEqnlst);
        }

        private void ProcessMeasPinCalcEqn(HardIpPattern pattern, int i, List<string> calcEqnlst, List<MeasPin> tmpMeasPins)
        {
            MeasPin measPin = pattern.MeasPins[i];
            if (!string.IsNullOrEmpty(measPin.CusStr))
            {
                if (!_allCusStr.ContainsKey(measPin.CusStr))
                {
                    _allCusStr.Add(measPin.CusStr, new List<string>());
                }

                _allCusStr[measPin.CusStr].Add(measPin.MeasType);
            }
            if (measPin.MeasType == MeasType.MeasLimit && !string.IsNullOrEmpty(measPin.CalcEqn) &&
                !calcEqnlst.Contains(measPin.CalcEqn))
            {
                tmpMeasPins.Add(measPin);
                calcEqnlst.AddRange(JudgeCalcEqnByJob(measPin.CalcEqn));
            }
            else if (measPin.MeasType == MeasType.MeasCalc && !string.IsNullOrEmpty(measPin.CalcEqn))
            {
                string forceCalcType = GetForceCalcType(measPin.MiscInfo);
                string calcType = !string.IsNullOrEmpty(forceCalcType) ? forceCalcType : GetMeasTypeForCalEqn(pattern.MeasPins.GetRange(0, i + 1), measPin.CalcEqn, measPin.RowNum);

                if (calcType.StartsWith("C") && Regex.IsMatch(measPin.CalcEqn, @"\((?<calc>.*)\)", RegexOptions.IgnoreCase))
                {
                    if (Regex.IsMatch(measPin.CalcEqn, RegPowerNotation, RegexOptions.IgnoreCase))
                    {
                        MatchCollection matches = Regex.Matches(measPin.CalcEqn, RegPowerNotation);
                        foreach (Match match in matches)
                        {
                            string value = Regex.Match(measPin.CalcEqn, RegPowerNotation, RegexOptions.IgnoreCase).Groups["value"].Value;
                            string power = Regex.Match(measPin.CalcEqn, RegPowerNotation, RegexOptions.IgnoreCase).Groups["power"].Value;
                            double ret = Math.Pow(double.Parse(value), double.Parse(power));
                            measPin.CalcEqn = measPin.CalcEqn.Replace(match.ToString(), ret.ToString());
                        }
                    }
                    if (Regex.IsMatch(measPin.CalcEqn, RegScientificNotation, RegexOptions.IgnoreCase))
                    {
                        MatchCollection matches = Regex.Matches(measPin.CalcEqn, RegScientificNotation, RegexOptions.IgnoreCase);
                        foreach (Match match in matches)
                        {
                            decimal realVal = decimal.Parse(match.ToString().Replace("(", "").Replace(")", ""), NumberStyles.Any);
                            measPin.CalcEqn = measPin.CalcEqn.Replace(match.ToString(), "(" + realVal + ")");
                        }
                    }
                    else
                    {
                        measPin.CalcEqn = measPin.CalcEqn;
                    }
                }
                List<string> eqns;
                if (string.IsNullOrEmpty(measPin.SkipUnit))
                {
                    eqns = ConvertPwrpinToDcSpec(JudgeCalcEqnByJob(calcType + ":" + measPin.CalcEqn + ":" + measPin.CusStr));
                }
                else
                {
                    eqns = ConvertPwrpinToDcSpec(JudgeCalcEqnByJob(calcType + "," + measPin.SkipUnit + ":" + measPin.CalcEqn + ":" + measPin.CusStr));
                }
                foreach (string eqn in eqns)
                {
                    MeasPin newEqn = measPin.Copy();
                    newEqn.PinName = eqn.Contains('=') ? eqn.Split('=')[0] + "=" + measPin.PinName : measPin.PinName;
                    tmpMeasPins.Add(newEqn);
                }

                calcEqnlst.AddRange(eqns);
            }
            else
            {
                tmpMeasPins.Add(measPin);
            }
        }

        private string GetForceCalcType(string miscInfo)
        {
            string calcType = "";
            foreach (string misc in miscInfo.Split(';'))
            {
                if (Regex.IsMatch(misc, "^forcecalctype:", RegexOptions.IgnoreCase))
                {
                    calcType = misc.Split(':')[1];
                    break;
                }
            }
            return calcType;
        }

        private string GetMeasTypeForCalEqn(List<MeasPin> measPins, string calcEqn, int row)
        {
            string type = "C";
            bool isFound = false;
            for (int i = measPins.Count - 1; i >= 0; i--)
            {
                if (measPins[i].CusStr != "" && calcEqn.ContainsIgnoreCase(measPins[i].CusStr.ToLower()))
                {
                    isFound = true;
                    if (measPins[i].MeasType != MeasType.MeasCalc && measPins[i].MeasType != MeasType.MeasLimit)
                    {
                        type = measPins[i].MeasType[4].ToString();
                        break;
                    }

                    //if its calc type, get inherit type (Workaround)
                    //TODO: 5/31 Discuss with CC & central to determine add new calc type
                    if (measPins[i].MeasType == MeasType.MeasCalc)
                    {
                        for (int j = measPins.Count - 1; j >= 0; j--)
                        {
                            if (measPins[j].CusStr != "" && measPins[i].CalcEqn.ContainsIgnoreCase(measPins[j].CusStr.ToLower()))
                            {
                                if (measPins[j].MeasType != MeasType.MeasCalc && measPins[j].MeasType != MeasType.MeasLimit)
                                {
                                    type = measPins[j].MeasType[4].ToString();
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            if (!isFound)
            {
                KeyValuePair<string, List<string>> measpinpat = _allCusStr.FirstOrDefault(clac => calcEqn.ContainsIgnoreCase(clac.Key.ToLower()));
                if (measpinpat.Value != null)
                {
                    isFound = true;
                    type = measpinpat.Value.Last().Equals("CALC", StringComparison.CurrentCultureIgnoreCase) ?
                        "C" : measpinpat.Value.Last()[4].ToString();
                }
            }
            if (!isFound)
            {
                string errorMessage = "CusStr of" + "\"" + calcEqn + "\" is not defined in this test item, and default calc type would be generated with C";
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongMeasContent_02,
                    _planSheet.SheetName,
                    row,
                    _planSheet.PlanHeaderIdx["measIndex"],
                    [calcEqn]
                );
                return "C";
            }
            return type;
        }

        internal string GetCalcEqnForPattern(string pinMiscInfo)
        {
            if (Regex.IsMatch(pinMiscInfo, HardIpConstData.Calc + ":", RegexOptions.IgnoreCase))
            {
                string calcEqnInPin = SearchInfo.GetCalculation(pinMiscInfo);
                return calcEqnInPin;
            }
            return "";
        }

        private static List<string> ConvertPwrpinToDcSpec(List<string> calcEqnList)
        {
            var result = new List<string>();
            var usedPin = new List<string>();
            foreach (string value in calcEqnList)
            {
                string regCalcEqn = @"(?<label>\w+)[^\(\w+\)]";
                MatchCollection matchPins = Regex.Matches(value.Split(':')[2], regCalcEqn, RegexOptions.IgnoreCase);
                string calcEqnVar = value;

                foreach (object matchPin in matchPins)
                {
                    if (int.TryParse(matchPin.ToString(), out int _))
                    {
                        continue;
                    }

                    string pinName =
                        Regex.Match(matchPin.ToString(), regCalcEqn, RegexOptions.IgnoreCase).Groups["label"].Value;
                    if (!pinName.Equals(value.Split(':').Last(), StringComparison.CurrentCultureIgnoreCase) &&
                        pinName.Contains("VDD", StringComparison.OrdinalIgnoreCase) && !usedPin.Contains(pinName))
                    {
                        usedPin.Add(pinName);
                        calcEqnVar = calcEqnVar.Replace(pinName, DataConvertor.ConvertValueWithGlbSpec(pinName));
                    }
                }
                result.Add(calcEqnVar);
            }

            return result;
        }

        private List<string> JudgeCalcEqnByJob(string calcExp)
        {
            var result = new List<string> { calcExp };
            if (Regex.IsMatch(calcExp, "alg:", RegexOptions.IgnoreCase))
            {
                return result;
            }
            string regCalcEqn = @"[\+\-\*\/]*\s*(?<label>(\w+|\-)+)(\(\w+\))*";
            MatchCollection matchPins = Regex.Matches(calcExp, regCalcEqn, RegexOptions.IgnoreCase);

            var cpTable = new Dictionary<string, string>();
            var ftTable = new Dictionary<string, string>();
            foreach (object matchPin in matchPins)
            {
                if (int.TryParse(matchPin.ToString(), out int _))
                {
                    continue;
                }

                string cpNetName = "";
                string ftNetName = "";
                string pinName = Regex.Match(matchPin.ToString(), regCalcEqn, RegexOptions.IgnoreCase).Groups["label"].Value;
                if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(pinName))
                {
                    var cpList = new List<string>();
                    var ftList = new List<string>();
                    foreach (string singlePin in TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(pinName))
                    {
                        cpNetName = "";
                        ftNetName = "";
                        TestPlanStatic.PowerMergeSheet.PowerMerge.GetCpFtNetName(singlePin, ref cpNetName, ref ftNetName);
                        if (cpNetName != "N/A" && !cpList.Contains(cpNetName))
                        {
                            cpList.Add(cpNetName);
                        }

                        if (ftNetName != "N/A" && !ftList.Contains(ftNetName))
                        {
                            ftList.Add(ftNetName);
                        }
                    }
                    cpList.Sort();
                    ftList.Sort();
                    if (cpList.SequenceEqual(ftList))
                    {
                        cpNetName = pinName;
                        ftNetName = pinName;
                    }
                    else
                    {
                        cpNetName = cpList.Count > 1 ? string.Join(",", cpList) : cpList[0];
                        ftNetName = ftList.Count > 1 ? string.Join(",", ftList) : ftList[0];
                    }
                }
                else
                {
                    TestPlanStatic.PowerMergeSheet.PowerMerge.GetCpFtNetName(pinName, ref cpNetName, ref ftNetName);
                    if (cpNetName == "N/A")
                    {
                        cpNetName = "";
                    }

                    if (ftNetName == "N/A")
                    {
                        ftNetName = "";
                    }
                }
                try
                {
                    if (cpNetName != ftNetName)
                    {
                        if (!cpTable.ContainsKey(pinName))
                        {
                            cpTable.Add(pinName, cpNetName);
                        }

                        if (!ftTable.ContainsKey(pinName))
                        {
                            ftTable.Add(pinName, ftNetName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                }
            }

            result = BuildCalcEqnByJobResult(calcExp, cpTable, ftTable, result);


            return result;
        }

        private List<string> BuildCalcEqnByJobResult(string calcExp, Dictionary<string, string> cpTable, Dictionary<string, string> ftTable, List<string> result)
        {
            if (cpTable.Count > 0 || ftTable.Count > 0)
            {
                result = new List<string>();
                if (cpTable.Count > 0)
                {
                    foreach (KeyValuePair<string, string> cpKey in cpTable)
                    {
                        foreach (string subKey in cpKey.Value.Split(','))
                        {
                            string eqn = GetFilterCalcEqn(calcExp, cpKey.Key, subKey);
                            if (!string.IsNullOrEmpty(eqn))
                            {
                                result.Add("CP=" + eqn);
                            }
                        }
                    }

                }
                if (ftTable.Count > 0)
                {
                    foreach (KeyValuePair<string, string> ftKey in ftTable)
                    {
                        foreach (string subKey in ftKey.Value.Split(','))
                        {
                            string eqn = GetFilterCalcEqn(calcExp, ftKey.Key, subKey);
                            if (!string.IsNullOrEmpty(eqn))
                            {
                                result.Add("FT=" + eqn);
                            }
                        }
                    }
                }
            }

            return result;
        }

        internal string GetFilterCalcEqn(string calcexp, string target, string replaceKey)
        {

            string result = calcexp;
            string regOper = @"[\+\-\*\/]*\s*(?<label>\w+)(\(\w+\))*";
            IEnumerable<Match> matches = Regex.Matches(calcexp, regOper, RegexOptions.IgnoreCase).Cast<Match>();
            foreach (Match match in matches)
            {
                if (target.Equals(match.Groups["label"].Value, StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrEmpty(replaceKey))
                    {
                        result = result.Replace(match.ToString(), "");
                    }
                    else
                    {
                        string newExp = Regex.Replace(match.Value, target, replaceKey);
                        result = result.Replace(match.ToString(), newExp);
                    }
                }
            }
            if (!Regex.Matches(calcexp, regOper, RegexOptions.IgnoreCase).Cast<Match>().Any())
            {
                return "";
            }

            return result;
        }
    }
}
