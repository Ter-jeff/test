using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Utility.HardIP;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class RegisterAssignmentChecker : HardIpPrecheckBase
    {
        private const string AssignRegStrWithKeyInterval = @"(?<Key>\w+)(\[|\:)*(?<Interval>\d+\:+\d+)\]*";

        private const string AssignRegStr2 = "^[0|1]+$";
        private const string AssignRegStr4 = @"^(?<label>[\w\-]+)$";
        private const string AssignRegStr5 = @"sweep\w*\((?<format>.*)\)";
        private const string AssignRegStr51 = @"sweep\w*\[(?<format>.*)\]";
        private const string AssignRegStr52 = @"sweep(?<order>\w*)\[(?<format>.*)\]@(?<alg>.*)";
        private const string AssignRegStrBin = "^(0b)?[01]+$";
        private const string AssignRegStrNonBin = "^(0x|0d)[0-9A-F]+";
        private const string AssignRegStrSelsram = @"selsram\(\)";
        private const string RegAssignLabelDefine = @"(?<label>[\w]+)([\(](?<value>[^)]*)[\)])[;]?";

        private readonly Dictionary<string, string> _dicLabel = new Dictionary<string, string>();
        private readonly List<string> _capList = new List<string>();
        private readonly Dictionary<string, string> _capListdir = new Dictionary<string, string>();

        public RegisterAssignmentChecker(HardIpSheet hardIpSheet, HardIpPattern pattern, Dictionary<string, string> planDic) : base(hardIpSheet, pattern)
        {
            _capListdir = planDic;
        }

        public override void Check()
        {
            int registerIndex = HardIpSheet.PlanHeaderIdx["registerIndex"];

            #region capture Name from MiscInfo

            _capList.AddRange(Pattern.MeasPins.Where(p => !string.IsNullOrEmpty(p.CusStr)).Select(p => p.CusStr));

            #endregion


            HardIpInfo patInfo = HardIpService.GetHardIpInfo(Pattern);
            List<string> sendBitNameList = patInfo.SendBitName.Split('+').ToList();
            List<string> sendBitStrlist = patInfo.SendBitStr.Split('+').ToList();
            //CheckPatternRegisterLimitation(patInfo, pattern);
            if (Pattern.Pattern.IsMultiple())
            {
                return;
            }

            if (string.IsNullOrEmpty(Pattern.RegisterAssignment) ^ string.IsNullOrEmpty(patInfo.SendBitStr))
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongRegisterAssignment_01,
                    Pattern.SheetName,
                    Pattern.RowNum,
                    registerIndex,
                    []
                );

            }

            if (Pattern.RegisterAssignment == "")
            {
                return;
            }

            string assignmentStr = Pattern.RegisterAssignment.Replace(" ", "");
            HandleRegAssignLabels(ref assignmentStr, _dicLabel);

            #region make register/bit size dictionary
            Dictionary<string, int> sendBitList = BuildSendBitList(sendBitNameList, sendBitStrlist);
            #endregion

            var wrongMessage = new List<string>();
            string nameStr = "";

            var missingKey = new List<string>();
            var currentBit = new List<string>();

            foreach (string assign in assignmentStr.Split(';'))
            {
                bool checkSendBitCnt = true;
                #region Checker digital source
                int sendBitCount = 0;
                if (!assign.Contains("="))
                {
                    if (_dicLabel.ContainsKey(assign)) // trimming VBT case
                    {
                    }
                }
                else
                {
                    string dataStr = assign.Split('=')[1];
                    nameStr = assign.Split('=')[0];
                    //syntax => 1. sweep(0,31,1) or sweep[1,3,1,5,1] or or sweep[1,3,1,5,1]@alg
                    if (Regex.IsMatch(dataStr, AssignRegStr5, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(dataStr, AssignRegStr51, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(dataStr, AssignRegStr52, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(dataStr, AssignRegStrSelsram, RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    foreach (string eachAssign in dataStr.Split('&'))
                    {
                        bool runMismatchCheck = ProcessEachAssign(eachAssign, assign, nameStr, patInfo, registerIndex, wrongMessage, missingKey, currentBit, ref checkSendBitCnt, ref sendBitCount);
                        if (!runMismatchCheck)
                        {
                            continue;
                        }

                        //dataStr Change to eachAssign due to support operand "&"
                        if (checkSendBitCnt && sendBitList.ContainsKey(nameStr) && sendBitList[nameStr] != sendBitCount)
                        {
                            string errorMessage =
                                $"Send bit: \"{nameStr}\", Mismatch bit width between pattern send bit: {sendBitList[nameStr]} and Register Assignment in test plan: {sendBitCount}.";
                            ErrorReportManager.AddError(
                                HardIpErrorType.E_WrongRegisterAssignment_02,
                                Pattern.SheetName,
                                Pattern.RowNum,
                                registerIndex,
                                [nameStr, $"{sendBitList[nameStr]}", $"{sendBitCount}"]
                            );
                        }
                    }
                }
                #endregion
            }
            if (missingKey.Count > 0)
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongRegisterAssignment_04,
                    Pattern.SheetName,
                    Pattern.RowNum,
                    registerIndex,
                    [string.Join(",", missingKey)]
                );
            }

            if (wrongMessage.Count > 0)
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongRegisterAssignment_05,
                    Pattern.SheetName,
                    Pattern.RowNum,
                    registerIndex,
                    [string.Join(";", wrongMessage)]
                );
            }

            #region Check sweepCode

            foreach (KeyValuePair<string, List<SweepCode>> sweepCodeSet in Pattern.SweepCodes)
            {
                CheckSweepCode(Pattern, sweepCodeSet.Value, registerIndex);
            }

            #endregion


        }

        private static Dictionary<string, int> BuildSendBitList(List<string> sendBitNameList, List<string> sendBitStrlist)
        {
            var sendBitList = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (sendBitNameList.Count == sendBitStrlist.Count)
            {
                for (int i = 0; i < sendBitNameList.Count; i++)
                {
                    if (!sendBitList.ContainsKey(sendBitNameList[i]))
                    {
                        if (int.TryParse(sendBitStrlist[i].Split('_')[sendBitStrlist[i].Split('_').Length - 1],
                                out int sendWidth))
                        {
                            sendBitList.Add(sendBitNameList[i], sendWidth);
                        }
                    }
                }
            }
            return sendBitList;
        }

        internal bool ProcessEachAssign(string eachAssign, string assign, string nameStr, HardIpInfo patInfo, int registerIndex, List<string> wrongMessage, List<string> missingKey, List<string> currentBit, ref bool checkSendBitCnt, ref int sendBitCount)
        {
            if (eachAssign.Contains("Reg_assign:"))
            {
                return false;
            }

            if (Regex.IsMatch(eachAssign, AssignRegStrBin, RegexOptions.IgnoreCase))
            {
                #region 0011000
                currentBit.Add(eachAssign);
                sendBitCount += eachAssign.Length;
                #endregion
            }
            else if (Regex.IsMatch(eachAssign, AssignRegStrNonBin, RegexOptions.IgnoreCase))
            {
                #region 0X00A0 or 0d0001
                string converted = RegisterConvertor.ConvertNumberToSrc(nameStr, eachAssign, patInfo);
                int infoBitLength = RegisterConvertor.GetSendBitLength(nameStr, patInfo);
                if (converted.Length != infoBitLength)
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongRegisterAssignment_02,
                        Pattern.SheetName,
                        Pattern.RowNum,
                        registerIndex,
                        [nameStr, $"{infoBitLength}", converted]
                    );
                }
                if (currentBit.Count > 0)
                {
                    wrongMessage.Add("\"" + eachAssign +
                                     "\" can not recongnize bit size because already exist data ");
                }

                currentBit.Add(eachAssign);
                return false;
                #endregion
            }
            else if (Regex.IsMatch(eachAssign, AssignRegStr4, RegexOptions.IgnoreCase))
            {
                #region check format "key"

                string[] array = eachAssign.Split(':');
                if (!Regex.IsMatch(array[0], AssignRegStr2, RegexOptions.IgnoreCase) &&
                    !_capList.Exists(p => p.Equals(array[0], StringComparison.OrdinalIgnoreCase)) &&
                    !_capListdir.Keys.ToList().Exists(p => p.Equals(array[0], StringComparison.OrdinalIgnoreCase)))
                {
                    if (!missingKey.Contains(array[0]))
                    {
                        missingKey.Add(array[0]);
                    }
                }
                checkSendBitCnt = false;
                return false;

                #endregion
            }
            else if (Regex.IsMatch(eachAssign, AssignRegStrWithKeyInterval, RegexOptions.IgnoreCase))
            {
                #region check format "key:2:0"

                Match match = Regex.Match(eachAssign, AssignRegStrWithKeyInterval, RegexOptions.IgnoreCase);
                string key = match.Groups["Key"].Value;
                if (!_capList.Exists(p => p.Equals(key, StringComparison.OrdinalIgnoreCase))

                    &&
                    !_capListdir.Keys.ToList().Exists(p => p.Equals(key, StringComparison.OrdinalIgnoreCase)))

                {
                    if (!missingKey.Contains(key))
                    {
                        missingKey.Add(key);
                    }
                }
                string strStart = match.Groups["Interval"].Value;
                if (int.TryParse(strStart.Split(':')[0], out int start) && int.TryParse(strStart.Split(':')[1], out int end))
                {
                    sendBitCount += Math.Abs(end - start) + 1;
                    if (start > end)
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.W_WrongRegisterAssignment_01,
                            Pattern.SheetName,
                            Pattern.RowNum,
                            registerIndex,
                            [$"{start}", $"{end}"]
                        );
                    }
                }
                else
                {
                    wrongMessage.Add("wrong format of start/stop \"" + eachAssign + "\"");
                }

                #endregion
            }
            else
            {
                wrongMessage.Add(", unrecognised format \"" + assign + "\"");
            }
            return true;
        }

        internal void CheckSweepCode(HardIpPattern planItem, List<SweepCode> sweepCodeSet, int registerColumnIndex)
        {
            string wrongMessage = "";
            bool wrongFlag = false;
            if (sweepCodeSet.Count > 1)
            {
                int steps = 0;
                if (sweepCodeSet.Any(p => p.Type == SweepCode.SweepType.Common) &&
                    sweepCodeSet.Any(p => p.Type == SweepCode.SweepType.Custom))
                {

                }
                foreach (SweepCode sweepItem in sweepCodeSet)
                {
                    if (!string.IsNullOrEmpty(sweepItem.StepMisc))
                    {
                        continue;
                    }

                    if (steps == 0)
                    {
                        steps = (sweepItem.End - sweepItem.Start) / sweepItem.Step;
                    }
                    else if ((sweepItem.End - sweepItem.Start) / sweepItem.Step != steps)
                    {
                        wrongFlag = true;
                        wrongMessage += ", SweepCodes' steps are inconsistent";
                        break;
                    }
                }
            }
            if (wrongFlag)
            {
                string errorMessage = "There is a Sweep format issue : " + wrongMessage;
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongRegisterAssignment_06,
                    planItem.SheetName,
                    planItem.RowNum,
                    registerColumnIndex,
                    [wrongMessage]
                );
            }
        }

        internal void HandleRegAssignLabels(ref string assignStr, Dictionary<string, string> dicLabel)
        {
            MatchCollection matches = Regex.Matches(assignStr, RegAssignLabelDefine, RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                string label = match.Groups["label"].ToString();
                if (label.ToLower() != "sweep" && label.ToLower() != "selsram")
                {
                    string value = match.Groups["value"].ToString();
                    if (!dicLabel.ContainsKey(label))
                    {
                        dicLabel.Add(label, value);
                    }

                    assignStr = assignStr.Replace(match.Value.Trim(';'), value);
                }
            }
        }
    }
}
