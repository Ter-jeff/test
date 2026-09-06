using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    internal class ForceConditionResolver
    {
        //eg: label1(IO1:V:0;IO2:V:0.4)
        private const string RegForceConditionLabelDefine = @"(?<label>[\w]+)[\s]*([\(](?<value>[^)]*)[\)])";

        //sweep(PinA:V:0.1) or sweepY(PinA:V:0.1)
        private static readonly Regex _regSweepVoltage =
            new Regex(@"sweep(?<label>\w*)\s*\((?<SweepStr>.*)\)(?<CustomSetting>:.+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        //NestSweep(PinA:V:0.1) or sweepY(PinA:V:0.1)
        private static readonly Regex _regNestSweepVoltage =
            new Regex(@"nestsweep(?<label>\D*)(?<order>\d*)\s*\((?<SweepStr>.*)\)(?<CustomSetting>:.+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        //NestSweep[0.001,0.002] or sweepY[0.001,0.002]
        private static readonly Regex _regNestSweepVoltageList =
            new Regex(@"nestsweep(?<label>\D*)(?<order>\d*)\s*\[(?<SweepStr>[^\]]+)\](?<CustomSetting>:.+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private readonly TestPlanSheet _planSheet;
        private readonly CalcEqnResolver _calcEqnResolver;
        private Dictionary<string, string> _forceConditionLabels = new Dictionary<string, string>();

        public ForceConditionResolver(TestPlanSheet planSheet, CalcEqnResolver calcEqnResolver)
        {
            _planSheet = planSheet;
            _calcEqnResolver = calcEqnResolver;
        }

        /// <summary>
        /// Resets the force condition label lookup. Must be called once per ConvertTpPatterns pass,
        /// mirroring the reset that used to happen on TestPlanPatParser's own field.
        /// </summary>
        internal void ResetForceConditionLabels()
        {
            _forceConditionLabels = new Dictionary<string, string>();
        }

        internal void ConvertForceLableForPatternNew(PatternRow patternRow, HardIpPattern pattern)
        {
            if (patternRow == null)
            {
                return;
            }

            //Handle interpose pre pattern force condition
            patternRow.ForceCondition.ForceCondition = ConvertForceLabelToValue(patternRow.ForceCondition.ForceCondition);

            //Hanld interpose pre measure force condition
            foreach (PatChildRow patChildRow in patternRow.PatChildRows)
            {
                foreach (TestPlanRow tprow in ((PatSubChildRow)patChildRow).TpRows)
                {
                    tprow.ForceCondition = ConvertForceLabelToValue(tprow.ForceCondition);
                }
            }

            //Handle post pattern force condition
            string miscInfo = "";
            if (patternRow == null)
            {
                miscInfo = pattern.MiscInfo;
            }
            patternRow.PostPatForceCondition = UpdatePostPatCalcEqn(ConvertForceLabelToValue(patternRow.PostPatForceCondition), miscInfo);
        }

        private string UpdatePostPatCalcEqn(string postPatForceCondition, string miscInfo)
        {
            var newPostPatForceConditionList = postPatForceCondition.Split(';').Where(x => !(x.Replace(" ", "").StartsWith("Calc:", StringComparison.OrdinalIgnoreCase) || x.Replace(" ", "").StartsWith("CalcArg:", StringComparison.OrdinalIgnoreCase))).ToList();
            string calcEqn = _calcEqnResolver.GetCalcEqnForPattern(postPatForceCondition);
            string postPatCalcEqn = calcEqn == "" ? _calcEqnResolver.GetCalcEqnForPattern(miscInfo) : calcEqn;
            if (!string.IsNullOrEmpty(postPatCalcEqn))
            {
                newPostPatForceConditionList.Add(postPatCalcEqn);
            }

            return string.Join(";", newPostPatForceConditionList);
        }

        private string ConvertForceLabelToValue(string forceStr)
        {
            //Excluing syntax like XShmoo(PinName:Level:From,To,Step:ShmooType:ShmooAlgorithm,[jump_step])
            if (HardIpConstData.RegShmoo.IsMatch(forceStr) || HardIpConstData.RegVtShmoo.IsMatch(forceStr))
            {
                return forceStr;
            }

            if (_regSweepVoltage.IsMatch(forceStr) || _regNestSweepVoltage.IsMatch(forceStr))
            {
                return forceStr;
            }

            //find all the label defination, store the value and replace it.
            //eg: label1(IO1:V:1;IO2:V:2), store "[label1, IO1:V:1;IO2:V:2]"
            //and replace label1(IO1:V:1;IO2:V:2) with "IO1:V:1;IO2:V:2"
            MatchCollection matches = Regex.Matches(forceStr, RegForceConditionLabelDefine, RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                string labelName = match.Groups["label"].ToString();
                string labelValue = match.Groups["value"].ToString();
                _forceConditionLabels[labelName] = labelValue;
                forceStr = forceStr.Replace(match.Value, labelValue);
            }
            forceStr = forceStr.Trim(';');

            //replace forceCondition label to actual value
            //eg: label1
            //should replace label1 with "IO1:V:1;IO2:V:2"
            string newStr = "";
            string[] splitStrArr = forceStr.Split(';');
            for (int i = 0; i < splitStrArr.Length; i++)
            {
                if (_forceConditionLabels.ContainsKey(splitStrArr[i]))
                {
                    newStr += _forceConditionLabels[splitStrArr[i]] + ";";
                }
                else
                {
                    newStr += splitStrArr[i] + ";";
                }
            }
            return newStr.Trim(';');
        }

        internal List<ForceCondition> DivideForceCondition(string forceStr, int rowNum)
        {
            var forceConditionlst = new List<ForceCondition>();

            //to support differential pin in force condition, we should replace "::" to "&" first
            List<string> forcelst = Regex.Replace(forceStr, "::", "&").Split(';').ToList();

            forcelst.RemoveAll(string.IsNullOrEmpty);
            if (forcelst.Count == 0)
            {
                return forceConditionlst;
            }

            //Support new format for Vdiff2 pin's forceCondition
            forcelst = ConvertVdiff2ForceFormat(forcelst);

            //Check format
            CheckForceStrFormat(forcelst, rowNum);

            //Convert force condition str to struct
            List<ForcePin> forcePinlst = ConvertForceStrToStruct(forcelst);

            //Seperate force condition, like USB_DP:V:5,6 to USB_DP:V:5 and USB_DP:V:6
            List<ForceCondition> separated = SeparateForceCondition(forcePinlst);

            return separated;
        }

        /// <summary>
        /// Seperate force condition, eg: "USB_DP:V:5,6;USB_DN:V:5" should seperate to two force conditons
        /// "USB_DP:V:5;USB_DN:V:5" and "USB_DP:V:6;USB_DN:V:5"
        /// </summary>
        /// <param name="forcePinlst"></param>
        /// <returns></returns>
        private List<ForceCondition> SeparateForceCondition(List<ForcePin> forcePinlst)
        {
            var forceConditionlst = new List<ForceCondition>();
            if (forcePinlst.Count == 0)
            {
                return forceConditionlst;
            }

            //find force condition count
            int count = 1;
            foreach (ForcePin forcePin in forcePinlst)
            {
                if (!Regex.IsMatch(forcePin.ForceType, "TERM|Sweep|SETUPFV|SETUPFI", RegexOptions.IgnoreCase)) // "MTR_TD_E:SETUPFI:6,0.00003 | MTR_TD_E:SETUPFV:6,0.00003"
                {
                    int forceCount = forcePin.ForceValue.Split(',').Length;
                    count = forceCount > count ? forceCount : count;
                    forcePin.ForceCnt = forceCount;
                }
            }

            //seperate to different force conditions
            if (count == 1)
            {
                var forceCondition = new ForceCondition();
                forceCondition.ForcePins.AddRange(forcePinlst);
                //forceCondition.ForcePins.Sort((x, y) => String.Compare(x.PinName, y.PinName));
                forceConditionlst.Add(forceCondition);
                return forceConditionlst;
            }

            foreach (ForcePin forcePin in forcePinlst)
            {
                var forceCondition = new ForceCondition();
                for (int i = 0; i < count; i++)
                {
                    ForcePin newForcePin = forcePin.Copy();
                    newForcePin.ForceValue = forcePin.ForceValue.Split(',').Length > i ?
                        forcePin.ForceValue.Split(',')[i] : forcePin.ForceValue.Split(',')[0];
                    forceCondition.ForcePins.Add(newForcePin);
                }
                forceConditionlst.Add(forceCondition);
                forceCondition.ForcePins.Sort((x, y) => string.Compare(x.PinName, y.PinName));
            }

            return forceConditionlst;
        }

        /// <summary>
        /// the format must be "XXX:XXX:[job]? or XXX:XXX:XXX:[job]?
        /// </summary>
        /// <param name="forcelst"></param>
        /// <param name="rowNum"></param>
        private void CheckForceStrFormat(List<string> forcelst, int rowNum)
        {
            foreach (string forceStr in forcelst)
            {
                if (HardIpConstData.RegShmoo.IsMatch(forceStr) ||
                    HardIpConstData.RegVtShmoo.IsMatch(forceStr) ||
                    _regSweepVoltage.IsMatch(forceStr) ||
                    _regNestSweepVoltage.IsMatch(forceStr) ||
                    HardIpConstData.RegMergeIndex.IsMatch(forceStr))
                {
                    continue;
                }

                List<string> forceArr = forceStr.Split(':').ToList();
                if (forceArr.Count < 2)
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongForceCondition_11,
                        _planSheet.SheetName,
                        rowNum,
                        _planSheet.ForceIndex,
                        []
                    );
                    continue;
                }

                if (forceArr.Last().Equals("norestore", StringComparison.CurrentCultureIgnoreCase))
                {
                    forceArr.RemoveAt(forceArr.Count - 1);
                }

                if (LocalSpecs.AllJobs.Exists(s => s.StartsWith(forceArr[forceArr.Count - 1], StringComparison.CurrentCultureIgnoreCase)))
                {
                    forceArr.RemoveAt(forceArr.Count - 1);
                }

                if (!(forceArr.Count == 2 || forceArr.Count == 3))
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongForceCondition_11,
                        _planSheet.SheetName,
                        rowNum,
                        _planSheet.ForceIndex,
                        []
                    );
                }
            }
        }

        private List<ForcePin> ConvertForceStrToStruct(List<string> forcelst)
        {
            var forcePinlst = new List<ForcePin>();
            string interpostMeas = "";
            if (forcelst.Count == 0)
            {
                return forcePinlst;
            }

            //Check interposeMeas
            var checkList = new List<string>(forcelst);
            foreach (string forceStr in checkList.Where(forceStr => forceStr.StartsWith("$")))
            {
                interpostMeas = forceStr.Replace("$", "");
                forcelst.Remove(forceStr);
            }

            forcelst = forcelst.Distinct().ToList();
            foreach (string forceStr in forcelst)
            {
                List<string> forceArr = forceStr.Split(':').ToList();

                if (HardIpConstData.RegShmoo.IsMatch(forceStr) || HardIpConstData.RegVtShmoo.IsMatch(forceStr) ||
                    _regSweepVoltage.IsMatch(forceStr) ||
                    _regNestSweepVoltage.IsMatch(forceStr) ||
                    _regNestSweepVoltageList.IsMatch(forceStr))
                {
                    continue;
                }

                if (forceArr.Count < 2)
                {
                    continue;
                }
                var forcePin = new ForcePin
                {
                    //Check norestore exists
                    IsRestore = !forceArr.LastOrDefault().Equals("norestore", StringComparison.CurrentCultureIgnoreCase)
                };

                if (!forcePin.IsRestore)
                {
                    forceArr.RemoveAt(forceArr.Count - 1);
                }
                //Read job form force str
                if (LocalSpecs.AllJobs.Exists(s => s.StartsWith(forceArr[forceArr.Count - 1], StringComparison.CurrentCultureIgnoreCase)))
                {
                    forcePin.ForceJob = forceArr[forceArr.Count - 1];
                    forceArr.RemoveAt(forceArr.Count - 1);
                }
                //Read force condition format type
                //pin:XXX:XXX => Normal, pin:XXX => Others
                forcePin.Type = forceArr.Count == 3 || forceArr.Count == 4 ? ForceConditionType.Normal : ForceConditionType.Others;
                //Read pin name and judge pin with power merge
                forcePin.PinName = string.Join(",", JudgePinByPowerMerge(Regex.Replace(forceArr[0], "&", "::").ToUpper()));
                //Read force value
                if (forcePin.Type == ForceConditionType.Normal)
                {
                    forcePin.ForceType = forceArr[1].ToUpper();
                    if (forceArr.Count > 2)
                    {
                        forcePin.ForceValue = string.Join(":", forceArr.Skip(2));
                    }
                }
                else
                {
                    forcePin.ForceValue = forceArr[1];
                    if (forceArr.Count != 2)
                    {
                        forcePin.ForceValue = forceStr.Substring(forceArr[0].Length + 1, forceStr.Length - forceArr[0].Length - 1);
                    }
                }
                //split pinName
                if (Regex.IsMatch(forcePin.PinName, "@"))
                {
                    string label = forcePin.PinName.Split('@')[0];
                    foreach (string voltage in label.Split(','))
                    {
                        forcePin.ForceLabelVoltages.Add(voltage);
                    }
                    forcePin.PinName = forcePin.PinName.Replace(label + "@", "");
                    if (forcePin.ForceLabelVoltages.Count == 3)
                    {
                        forcePin.ForceLabelVoltages.Clear();
                    }
                }
                //Add interposeMeas
                if (!string.IsNullOrEmpty(interpostMeas))
                {
                    forcePin.ForceInterPostMeas = interpostMeas;
                }
                foreach (string pin in forcePin.PinName.Split(','))
                {
                    ForcePin newPin = forcePin.Copy();
                    newPin.PinName = pin;
                    forcePinlst.Add(newPin);
                }

            }
            return forcePinlst;
        }

        /// <summary>
        /// if vdiff2 pin forceCondition like "PinName:Vn:0.3,0.2;PinName:Vp:0.3,0.2"
        /// change to "PinName:V1n:0.3;PinName:V2n:0.2;PinName:V1p:0.3;PinName:V2p:0.2"
        /// </summary>
        /// <param name="forcelst"></param>
        private List<string> ConvertVdiff2ForceFormat(List<string> forcelst)
        {
            var result = new List<string>();
            var vn = new List<string>();
            var vp = new List<string>();
            string pPin = "";
            string nPin = "";
            string regVn = @"\:v\dn\:";
            string regVp = @"\:v\dp\:";
            foreach (string force in forcelst.ToArray())
            {
                if (Regex.IsMatch(force, regVn, RegexOptions.IgnoreCase))
                {
                    nPin = force.Split(':')[0];
                    vn.Add(force.Split(':')[2]);
                }
                if (Regex.IsMatch(force, regVp, RegexOptions.IgnoreCase))
                {
                    pPin = force.Split(':')[0];
                    vp.Add(force.Split(':')[2]);
                }
            }

            if (vn.Count > 0)
            {
                result.Add($"{nPin}:V:{string.Join("&", vn)}");
            }

            if (vp.Count > 0)
            {
                result.Add($"{pPin}:V:{string.Join("&", vp)}");
            }

            if (result.Count == 0)
            {
                result.AddRange(forcelst);
            }

            return result;
        }

        internal List<string> JudgePinByPowerMerge(string pinName)
        {
            if (string.IsNullOrEmpty(pinName))
            {
                return new List<string> { pinName };
            }

            string cpNetName = "";
            string ftNetName = "";
            var result = new List<string>();

            if (TestPlanStatic.PowerMergeSheet == null)
            {
                return result;
            }

            TestPlanStatic.PowerMergeSheet.PowerMerge.GetCpFtNetName(pinName, ref cpNetName, ref ftNetName);
            if (cpNetName != ftNetName)
            {
                if (cpNetName != "")
                {
                    string cpPin = "CP=" + cpNetName.Replace(",", ",CP=");
                    result.Add(cpPin);
                }
                if (ftNetName != "")
                {
                    string ftPin = "FT=" + ftNetName.Replace(",", ",FT=");
                    result.Add(ftPin);
                }
            }
            else if (cpNetName != "")
            {
                result.Add(cpNetName);
            }
            else
            {
                result.Add(pinName);
            }

            return result;
        }
    }
}
