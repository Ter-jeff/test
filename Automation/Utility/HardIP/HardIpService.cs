using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

namespace Automation.Utility.HardIP
{
    public static class HardIpService
    {
        private static readonly Regex _regex1 = new Regex("(?<relayon>RelayOff:[^R]+).*");
        private static readonly Regex _regex2 = new Regex("(?<relayoff>RelayOn:[^R]+).*");
        private static readonly Regex _regex = new Regex(HardIpConstData.IgnorePatDigSrc, RegexOptions.IgnoreCase);
        private static readonly Regex _regex3 = new Regex(HardIpConstData.IgnorePatInfo, RegexOptions.IgnoreCase);

        public static string ReverseRelaySetting(string setting)
        {
            string relayOn = _regex1.Match(setting).Groups["relayon"].ToString().Replace("Off", "On");
            string relayOff = _regex2.Match(setting).Groups["relayoff"].ToString().Replace("On", "Off");
            return (relayOn + "_" + relayOff).Trim('_');
        }

        public static string GetFlowForLoopIntegerName(HardIpPattern pattern, HardIpInputData hardIpInputData, bool isCsUsing = false)
        {
            string flowForLoopIntegerName = "";
            if (pattern.Pattern.IsMultiple())
            {
                var sweepList = new List<string>();
                foreach (HardIpPattern pat in pattern.BurstPatterns)
                {
                    if (pat.SweepCodes.Count > 0)
                    {
                        foreach (KeyValuePair<string, List<SweepCode>> sweepItems in pat.SweepCodes)
                        {
                            var flowLoopNameInfo = new List<string>();
                            flowLoopNameInfo.AddRange(sweepItems.Value.Select(x => x.GetFlowForLoopInfo(isCsUsing)));
                            sweepList.Add($"SrcCodeIndx{sweepItems.Key}" + ";" + string.Join(";", flowLoopNameInfo));
                        }
                    }
                }
                flowForLoopIntegerName = string.Join("|", sweepList);
            }
            else
            {
                if (pattern.SweepCodes.Count > 0)
                {
                    var sweepList = new List<string>();
                    foreach (KeyValuePair<string, List<SweepCode>> sweepItems in pattern.SweepCodes)
                    {
                        var flowLoopNameInfo = new List<string>();
                        flowLoopNameInfo.AddRange(sweepItems.Value.Select(x => x.GetFlowForLoopInfo(isCsUsing)));
                        sweepList.Add($"SrcCodeIndx{sweepItems.Key}" + ";" + string.Join(";", flowLoopNameInfo));
                    }
                    flowForLoopIntegerName = string.Join("|", sweepList);
                }
            }

            if (flowForLoopIntegerName.Length >= 6000)
            {
                var registerAssignItem = new HardIpRegAssign
                {
                    SubBlockName = CommonGenerator.GetRegAssignName(pattern),
                    Type = RegisterAssignType.DigSrc_FlowForLoopIntegerName
                };
                if (!hardIpInputData.HardIpRegAssigns.Exists(x => x.SubBlockName.Equals(registerAssignItem.SubBlockName, StringComparison.CurrentCultureIgnoreCase) && x.Type == registerAssignItem.Type))
                {
                    var assignList = new List<List<string>>();
                    foreach (string item in flowForLoopIntegerName.Split('|').ToList())
                    {
                        var list = new List<string>();
                        int i = 0;

                        foreach (string info in item.Split(';').ToList())
                        {
                            if (i == 0)
                            {
                                list.Add(info);
                            }
                            else if (i == 1)
                            {
                                list.Add(info);
                                assignList.Add(list);
                            }
                            else
                            {
                                list = new List<string> { "", info };
                                assignList.Add(list);
                            }
                            i++;
                        }

                    }
                    registerAssignItem.RegAssignList = assignList;
                    hardIpInputData.HardIpRegAssigns.Add(registerAssignItem);
                }
                flowForLoopIntegerName = $"Reg_assign:{registerAssignItem.SubBlockName}";

            }

            return flowForLoopIntegerName.Trim(';');
        }

        public static HardIpInfo GetHardIpInfo(HardIpPattern pattern)
        {
            HardIpInfo hardIpInfo = new HardIpInfo();
            var payloadName = new List<string>();
            foreach (string patternName in pattern.Pattern.PatternSetList.SelectMany(p => p))
            {
                payloadName.Add(patternName);
                hardIpInfo.MergeInfo(GetHardIpInfo(patternName));
            }
            hardIpInfo.Payload = string.Join("+", payloadName);

            if (!string.IsNullOrEmpty(hardIpInfo.MultiDsscOut))
            {
                hardIpInfo.MultiDsscOut = hardIpInfo.MultiDsscOut.Substring(0, hardIpInfo.MultiDsscOut.Length - 1);
            }

            bool isIgnorePat = _regex3.IsMatch(pattern.MiscInfo);
            if (isIgnorePat && hardIpInfo.SeqInfo.Count > 0)
            {
                var newHardIpInfo = new HardIpInfo();
                bool ignorePatDigSrc = _regex.IsMatch(pattern.MiscInfo);
                if (ignorePatDigSrc)
                {
                    newHardIpInfo.CapBit = hardIpInfo.CapBit;
                    newHardIpInfo.CapBitName = hardIpInfo.CapBitName;
                    newHardIpInfo.CapBitStr = hardIpInfo.CapBitStr;
                    newHardIpInfo.CapPinName = hardIpInfo.CapPinName;
                    newHardIpInfo.DigSrcAssign = hardIpInfo.DigSrcAssign;
                    newHardIpInfo.DigSrcDataWidth = hardIpInfo.DigSrcDataWidth;
                    newHardIpInfo.DsscOut = hardIpInfo.DsscOut;
                    newHardIpInfo.MultiDsscOut = hardIpInfo.MultiDsscOut;
                    newHardIpInfo.Payload = hardIpInfo.Payload;
                    newHardIpInfo.Subroutine = hardIpInfo.Subroutine;
                    newHardIpInfo.TimeSet = hardIpInfo.TimeSet;
                    newHardIpInfo.IsIgnoreComment = true;
                    newHardIpInfo.PatInfoExist = hardIpInfo.PatInfoExist;
                    return newHardIpInfo;
                }

                newHardIpInfo.PatInfoExist = hardIpInfo.PatInfoExist;
                newHardIpInfo.CapBit = hardIpInfo.CapBit;
                newHardIpInfo.CapBitName = hardIpInfo.CapBitName;
                newHardIpInfo.CapBitStr = hardIpInfo.CapBitStr;
                newHardIpInfo.CapPinName = hardIpInfo.CapPinName;
                newHardIpInfo.CallSubrs = hardIpInfo.CallSubrs;
                newHardIpInfo.CallSubrsCnt = hardIpInfo.CallSubrsCnt;
                newHardIpInfo.DigSrcAssign = hardIpInfo.DigSrcAssign;
                newHardIpInfo.DigSrcDataWidth = hardIpInfo.DigSrcDataWidth;
                newHardIpInfo.DsscOut = hardIpInfo.DsscOut;
                newHardIpInfo.MultiDsscOut = hardIpInfo.MultiDsscOut;
                newHardIpInfo.Payload = hardIpInfo.Payload;
                newHardIpInfo.SendBit = hardIpInfo.SendBit;
                newHardIpInfo.SendBitName = hardIpInfo.SendBitName;
                newHardIpInfo.SendBitStr = hardIpInfo.SendBitStr;
                newHardIpInfo.SendPinName = hardIpInfo.SendPinName;
                newHardIpInfo.Subroutine = hardIpInfo.Subroutine;
                newHardIpInfo.TimeSet = hardIpInfo.TimeSet;
                newHardIpInfo.IsIgnoreComment = true;
                return newHardIpInfo;
            }
            if (_regex.IsMatch(pattern.MiscInfo))
            {
                hardIpInfo.SendBit = "0";
                hardIpInfo.SendBitName = "";
                hardIpInfo.SendBitStr = "";
                hardIpInfo.SendPinName = "";
            }
            return hardIpInfo;
        }

        public static HardIpInfo GetHardIpInfo(string patternName)
        {
            HardIpInfo info = LocalSpecs.HardIpInfos.GetHardIpInfo(patternName, useThisVersion: true);
            if (info == null)
            {
                return new HardIpInfo();
            }

            if (LocalSpecs.Options.Device == EnumDevice.LCD && info.CapBit > 200000)
            {
                string capBitStr = info.CapBitStr;
                int capBit = info.CapBit;
                info.CapBitName = info.CapBitName.Split('+')[0];
                info.CapBitStr = capBitStr.Split('+')[0].Split('_')[0] + "_" + capBit;
                if (!string.IsNullOrEmpty(info.DsscOut))
                {
                    string firstPart = info.DsscOut.Replace("DSSC_OUT,", "").Split(',')[0];
                    int colonIndex = firstPart.IndexOf(':');
                    if (colonIndex != -1)
                    {
                        string tNameWithoutBitSize = firstPart.Substring(colonIndex + 1);
                        info.DsscOut = $"DSSC_OUT,{capBit}:{tNameWithoutBitSize}";
                    }
                }
            }

            return info;
        }

        public static bool IdsNoFuse(string miscInfo)
        {
            bool math = false;

            math = miscInfo.Split(';').ToList().Exists(s => Regex.IsMatch(s, HardIpConstData.IdsNoFuse, RegexOptions.IgnoreCase));

            math = math || !miscInfo.Split(';').ToList().Exists(s => Regex.IsMatch(s, "EFUSEPOWER_Pin", RegexOptions.IgnoreCase));

            return math;
        }

        public static bool IsNoPattern(string patternName)
        {
            return patternName.Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsNandPattern(string patternName)
        {
            return Regex.IsMatch(patternName, HardIpConstData.RegNand, RegexOptions.IgnoreCase);
        }

        public static bool IsSpiPattern(string patternName)
        {
            return Regex.IsMatch(patternName, HardIpConstData.RegSpi, RegexOptions.IgnoreCase);
        }

        public static bool IsCz2Only(string miscInfo)
        {
            bool math = false;
            math = miscInfo.Split(';').ToList().Exists(s => s.Equals(HardIpConstData.Cz2Only, StringComparison.OrdinalIgnoreCase));
            return math;
        }

        public static bool IsNestSweepVoltage(HardIpPattern pattern)
        {
            if (pattern.SweepVoltage.Count > 0)
            {
                return pattern.SweepVoltage.All(x => x.Key.ContainsIgnoreCase("NESTSWEEP"));
            }
            return pattern.SweepVoltage.Count > 0;
        }

        public static string ActualLabelVoltage(string defaultVoltage, HardIpPattern pattern)
        {
            string voltage = defaultVoltage;
            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.HvOnly, RegexOptions.IgnoreCase))
            {
                voltage = HardIpConstData.LabelHv;
            }

            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.LvOnly, RegexOptions.IgnoreCase))
            {
                voltage = HardIpConstData.LabelLv;
            }

            if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.NvOnly, RegexOptions.IgnoreCase))
            {
                voltage = HardIpConstData.LabelNv;
            }

            return voltage;
        }

        public static bool PatFlowNoNeedToGen(string labelVoltage, string miscInfo)
        {
            switch (labelVoltage)
            {
                case HardIpConstData.LabelNv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RunHv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RunLv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RemoveNv, RegexOptions.IgnoreCase))
                    {
                        return true;
                    }
                    break;

                case HardIpConstData.LabelLv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RunHv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RunNv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RemoveLv, RegexOptions.IgnoreCase))
                    {
                        return true;
                    }
                    break;

                case HardIpConstData.LabelHv:
                    if (Regex.IsMatch(miscInfo, HardIpConstData.RunNv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RunLv, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(miscInfo, HardIpConstData.RemoveHv, RegexOptions.IgnoreCase))
                    {
                        return true;
                    }
                    break;
            }
            return false;
        }

        public static string GetOriginalTtrBranchItem(string miscInfo)
        {
            foreach (string assign in miscInfo.Split(';'))
            {
                string[] asignArr = assign.Split(':');
                if (asignArr.Length == 2 &&
                    asignArr[0].Equals(HardIpConstData.OriginalTtrBranch, StringComparison.OrdinalIgnoreCase))
                {
                    return asignArr[1];
                }
            }
            return "";
        }

        public static string GetMappingTtrBranchItem(string miscInfo)
        {
            foreach (string assign in miscInfo.Split(';'))
            {
                string[] asignArr = assign.Split(':');
                if (asignArr.Length == 2 &&
                    asignArr[0].Equals(HardIpConstData.MappingTtrBranch, StringComparison.OrdinalIgnoreCase))
                {
                    return asignArr[1];
                }
            }
            return "";
        }

        /// <summary>
        /// Get InstNameSubStr base on what specified in miscInfo: "InstNameSubStr:XXX"
        /// </summary>
        /// <returns></returns>
        public static string GetInstNameSubStr(string miscInfo)
        {
            string instNameSubStr = "";
            foreach (string info in miscInfo.Split(';'))
            {
                if (info.Contains(":"))
                {
                    string[] misArr = info.Split(':');
                    if (misArr.Length == 2 && misArr[0].Equals(HardIpConstData.InstNameSubStr, StringComparison.OrdinalIgnoreCase))
                    {
                        instNameSubStr = misArr[1];
                        break;
                    }
                }
            }
            return instNameSubStr;
        }

        public static string GetRepeatMapping(string miscInfo)
        {
            foreach (string param in miscInfo.Split(';'))
            {
                if (!param.Contains(":"))
                {
                    continue;
                }

                string paramName = param.Split(':')[0];
                string paramValue = param.Split(':')[1];
                if (paramName == HardIpConstData.Limit)
                {
                    return paramValue;
                }
            }
            return "";
        }

        public static string CalcuteLimit(string limitStr, string repeatStr)
        {
            if (limitStr == "" || limitStr == "0")
            {
                return limitStr;
            }

            const string regRate = @"(?<Num>\d+(\.\d){0,1})x+$";
            double limitValue = Convert.ToDouble(limitStr);
            double rate = 1;
            if (repeatStr != "")
            {
                rate = double.Parse(Regex.Match(repeatStr, regRate).Groups["Num"].ToString());
            }

            double result = limitValue * rate;
            return result.ToString();
        }

    }
}
