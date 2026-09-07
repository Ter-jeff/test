using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;

namespace Automation.GenerateIgxl.HardIp.InputReader
{
    public class RegisterAssignParser
    {
        // syntax-1 Reg1=0001 => 8
        // syntax-2 Reg1=0b110 is 6, reg2=0d10 is 10, reg3=0x0A is 10
        // syntax-3 Reg1=Cal_B:0:3 if Cal_B=010011 then Reg1=0100      => Handle in VBT
        // syntax-4 Reg1=cal_B&10, if cal_B=11100000 then 1110000010   => Handle in VBT
        // syntax-5 Reg1=(cal_A-2)&10 if cal_A=0100 then Reg1=000010   => Handle in VBT
        // syntax-6 trim_Reuse(trim=Cal_All:4:7)
        // syntax-7 Repeat=0010 is reg1=0010; reg2=0010; reg3=0010
        // syntax-8 repeat=rd0:0:3:copy:4                              => Handle in VBT
        // syntax-9 reg1=10;reg1=11;reg2=01;reg1=00 =>　reg1_1=10;reg1_2=11;reg2_1=01;reg1_3=00
        // syntax-10 reg1=sweep(0,3,1)
        // syntax-11 reg1=sweep(0,3,1,graycode)
        // syntax-12 reg1=sweep(0,3,1):bit_repeat:2 => 00 00,11 00,00 11,11 11
        // syntax-13 reg1=sweep1[00,11]
        // syntax-14 reg1=sweep1[00,11]@BinToGray => 00,11@BinToGray
        // syntax-15 reg1=NestSweep1[00,11]@BinToGray => 00,11@BinToGray
        // syntax-16 reg1=NestSweep1(0,3,1,graycode)

        private const string RegSpecialFormat = "^(0x|0b|0d)";
        private const string RegSweepCode = @"sweep(?<order>\w*)[(](?<format>[^)]+)[)]";
        private const string RegSweepCode1 = @"(?(?=.*@)sweep(?<order>\w*)\[(?<format>.*)\]@(?<alg>.*)|sweep(?<order>\w*)\[(?<format>.*)\])";
        private const string RegNestSweepCode = @"nestsweep(?<order>\w*)[(](?<format>[^)]+)[)]";
        //private const string RegRepeat = @"repeat=.*";

        private HardIpInfo _patInfo;
        private bool _ignorePatComment;

        private string _sheetName = "";
        private int _rowNum;
        private int _registerIndex;

        public void ParseRegisterAssign(Dictionary<string, HardIpSheet> planDic)
        {
            try
            {
                foreach (KeyValuePair<string, HardIpSheet> sheet in planDic)
                {
                    _sheetName = sheet.Key;
                    _registerIndex = sheet.Value.PlanHeaderIdx["registerIndex"];
                    foreach (HardIpPattern pattern in sheet.Value.Rows)
                    {
                        _rowNum = pattern.RowNum;

                        _patInfo = HardIpService.GetHardIpInfo(pattern);

                        _ignorePatComment = Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatInfo, RegexOptions.IgnoreCase);

                        List<string> patInfoSendBitNameList = !string.IsNullOrEmpty(_patInfo.SendBitName) ? _patInfo.SendBitName.Split('+', '|').Distinct().ToList() : new List<string>();

                        List<string> newAssignmentStrList = GetNewAssignmentStrList(pattern, patInfoSendBitNameList);

                        pattern.RegisterAssignment = string.Join("|", newAssignmentStrList);
                        pattern.DigSrcEquation = GetDigSrcEquation(pattern, _patInfo);
                        if (pattern.Pattern.IsMultiple())
                        {
                            int patIdx = 0;
                            foreach (HardIpPattern subPattern in pattern.BurstPatterns)
                            {
                                subPattern.RegisterAssignment = newAssignmentStrList[patIdx];
                                patIdx++;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Response.Report("Parsing register assignment information failed at " + _sheetName + ": Row " + _rowNum + "--" + e.StackTrace, EnumMessageLevel.Error, 0);
            }
        }

        private void CheckAssignmentLength(HardIpPattern lpattern, List<HardIpPattern> patterns)
        {
            int busrstAssignmentCount = lpattern.RegisterAssignment.Trim().Split('|').Length;
            if (busrstAssignmentCount == 1)
            {
                return;
            }

            if (patterns.Count != busrstAssignmentCount)
            {
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongRegisterAssignment_07,
                    lpattern.SheetName,
                    lpattern.RowNum,
                    _registerIndex,
                    [$"{busrstAssignmentCount}", $"{patterns.Count}"]
                );
            }
        }

        private List<string> GetNewAssignmentStrList(HardIpPattern lpattern, List<string> patInfoSendBitNameList)
        {
            List<string> newAssignmentStrList = new List<string>();
            List<HardIpPattern> patterns = lpattern.BurstPatterns.Count > 0 ? lpattern.BurstPatterns : new List<HardIpPattern> { lpattern };
            bool isOnlyOneAssignmentPart = lpattern.RegisterAssignment.Trim().Split('|').Length == 1;
            CheckAssignmentLength(lpattern, patterns);
            var planAllBitNameList = new List<string>();
            foreach (HardIpPattern pattern in patterns)
            {
                string newAssignmentStr = "";
                HardIpInfo patinfo = HardIpService.GetHardIpInfo(pattern.Pattern.GetPatternName());
                List<string> planBitNameList = new List<string>();
                foreach (string regItem in pattern.RegisterAssignment.Split(';', '|'))
                {
                    if (string.IsNullOrEmpty(regItem))
                    {
                        continue;
                    }

                    if (Regex.IsMatch(pattern.RegisterAssignment, "Reg_assign:", RegexOptions.IgnoreCase))
                    {
                        newAssignmentStr = regItem;
                        continue;
                    }
                    #region reg=0000 or reg=Sweep(0,15,1) or reg=rd0
                    if (regItem.Contains("="))
                    {
                        string sendBitName = regItem.Split('=')[0];
                        string sendValue = regItem.Split('=')[1];
                        planBitNameList.Add(sendBitName);

                        if (!patinfo.SendBitName.Split('+', '|').Any(x => x.Equals(sendBitName, StringComparison.OrdinalIgnoreCase)))
                        {
                            if (!isOnlyOneAssignmentPart)
                            {
                                ErrorReportManager.AddError(
                                    HardIpErrorType.E_WrongRegisterAssignment_08,
                                    lpattern.SheetName,
                                    lpattern.RowNum,
                                    _registerIndex,
                                    [sendBitName, pattern.Pattern.GetPatternName()]
                                );
                            }
                            continue;
                        }
                        newAssignmentStr += ParseSrcValue(sendBitName, sendValue, patinfo.IsRegisterReverse(sendBitName), pattern) + ";";
                        continue;
                    }
                    #endregion

                    #region 0000 or Sweep(0,15,1) or RD0
                    if (patInfoSendBitNameList.Count > planBitNameList.Distinct().ToList().Count)
                    {
                        string sendName = patInfoSendBitNameList[planBitNameList.Distinct().ToList().Count];
                        planBitNameList.Add(sendName);
                        newAssignmentStr += ParseSrcValue(sendName, regItem, patinfo.IsRegisterReverse(sendName), pattern) + ";";
                    }
                    #endregion
                }
                newAssignmentStrList.Add(newAssignmentStr.Trim(';'));

                IEnumerable<string> patInfoSendBitNameLower = patinfo.SendBitName.Split('+', '|').Select(x => x.ToLower());
                IEnumerable<string> planBitNameListLower = planBitNameList.Select(x => x.ToLower());
                IEnumerable<string> checkPlanNotContainsInfo = patInfoSendBitNameLower.Except(planBitNameListLower);

                foreach (string loseDefine in checkPlanNotContainsInfo.Where(x => !string.IsNullOrEmpty(x.Trim())))
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongRegisterAssignment_09,
                        lpattern.SheetName,
                        lpattern.RowNum,
                        _registerIndex,
                        [loseDefine, pattern.Pattern.GetPatternName()]
                    );
                }
                planAllBitNameList.AddRange(planBitNameList);
            }
            if (!_ignorePatComment)
            {
                if (patInfoSendBitNameList.Count != planAllBitNameList.Distinct().Count())
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongRegisterAssignment_10,
                        lpattern.SheetName,
                        lpattern.RowNum,
                        _registerIndex,
                        [$"{patInfoSendBitNameList.Count}", $"{planAllBitNameList.Distinct().Count()}"]
                    );
                }
                List<string> lowercaseBitNameList = new List<string>();
                foreach (string bitname in planAllBitNameList)
                {
                    lowercaseBitNameList.Add(bitname.ToLower());
                }

                IEnumerable<string> checkInfoNotContainsPlanReg = patInfoSendBitNameList.Except(lowercaseBitNameList);
                string.Join("\n", planAllBitNameList);
                if (checkInfoNotContainsPlanReg.Any())
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongRegisterAssignment_11,
                        lpattern.SheetName,
                        lpattern.RowNum,
                        _registerIndex,
                        [string.Join(",", checkInfoNotContainsPlanReg)]
                    );
                }

                IEnumerable<string> checkPlanNotContainsInfoReg = lowercaseBitNameList.Except(patInfoSendBitNameList);
                if (checkPlanNotContainsInfoReg.Any())
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongRegisterAssignment_12,
                        lpattern.SheetName,
                        lpattern.RowNum,
                        _registerIndex,
                        [string.Join(",", checkPlanNotContainsInfoReg)]
                    );
                }
            }
            return newAssignmentStrList;
        }

        private string GetDigSrcEquation(HardIpPattern pattern, HardIpInfo patInfo)
        {
            if (pattern.DigSrcEquation.Trim('|') != "")
            {
                return pattern.DigSrcEquation;
            }

            return patInfo.SendBitName;
        }

        internal string ParseSrcValue(string sendBitName, string sendValue, bool isNeedReverse, HardIpPattern pattern)
        {
            //deal with sweep code
            #region Process Sweep code
            if (Regex.IsMatch(sendValue, RegSweepCode, RegexOptions.IgnoreCase) ||
                Regex.IsMatch(sendValue, RegSweepCode1, RegexOptions.IgnoreCase))
            {
                return GetSweepCode(sendBitName, sendValue, pattern, _patInfo);
            }

            #endregion
            var tmpSendValue = new List<string>();
            #region Process Customize src
            if (Regex.IsMatch(sendValue, @"\[(?<value>.*)\]", RegexOptions.IgnoreCase))
            {
                string subValueStr =
                    Regex.Match(sendValue, @"\[(?<value>.*)\]", RegexOptions.IgnoreCase).Groups["value"].ToString();
                foreach (string subValue in subValueStr.Split(','))
                {
                    string tmpStr = subValue;
                    if (Regex.IsMatch(tmpStr, RegSpecialFormat, RegexOptions.IgnoreCase))
                    {
                        tmpStr = RegisterConvertor.ConvertNumberToSrc(sendBitName, tmpStr, _patInfo);
                    }

                    tmpSendValue.Add(tmpStr);
                }
                return sendBitName + "=[" + string.Join(",", tmpSendValue) + "]";
            }
            #endregion

            #region Convert value with non binary format
            foreach (string subValue in sendValue.Split('&'))
            {
                string tmpStr = subValue;
                //deal with special format 0x/0b/0d
                if (Regex.IsMatch(tmpStr, RegSpecialFormat, RegexOptions.IgnoreCase))
                {
                    tmpStr = RegisterConvertor.ConvertNumberToSrc(sendBitName, tmpStr, _patInfo);
                }

                if (isNeedReverse && int.TryParse(tmpStr, out int _))
                {
                    char[] cArray = tmpStr.ToCharArray();
                    Array.Reverse(cArray);
                    tmpStr = new string(cArray);
                }
                tmpSendValue.Add(tmpStr);
            }
            #endregion
            sendValue = string.Join("&", tmpSendValue);
            if (sendValue.Equals("0"))
            {
                return sendBitName + "=" + string.Join("", Enumerable.Repeat("0", RegisterConvertor.GetSendBitLength(sendBitName, _patInfo)));
            }

            return sendBitName + "=" + sendValue;
        }

        internal string GetSweepCode(string sendBitName, string sweepCode, HardIpPattern pattern, HardIpInfo patInfo)
        {
            int srclengthInPatInfo = RegisterConvertor.GetSendBitLength(sendBitName, patInfo);
            string format;
            string order;
            if (sweepCode.ToUpper().StartsWith("NEST"))
            {
                format = Regex.Match(sweepCode, RegNestSweepCode, RegexOptions.IgnoreCase).Groups["format"].ToString();
                order = Regex.Match(sweepCode, RegNestSweepCode, RegexOptions.IgnoreCase).Groups["order"].ToString().Trim();
            }
            else
            {
                format = Regex.Match(sweepCode, RegSweepCode, RegexOptions.IgnoreCase).Groups["format"].ToString();
                order = Regex.Match(sweepCode, RegSweepCode, RegexOptions.IgnoreCase).Groups["order"].ToString().Trim();
            }

            if (srclengthInPatInfo > 0)
            {
                var sCode = new SweepCode { SendBitName = sendBitName };
                if (Regex.IsMatch(sweepCode, RegSweepCode, RegexOptions.IgnoreCase) && !Regex.IsMatch(sweepCode, RegNestSweepCode, RegexOptions.IgnoreCase))
                {
                    sCode.Type = SweepCode.SweepType.Common;
                    sCode.SweepInfo = format;
                }
                else
                {
                    sCode.Type = SweepCode.SweepType.Custom;
                    string customCode = Regex.Match(sweepCode, RegSweepCode1, RegexOptions.IgnoreCase).Groups["format"].ToString();
                    if (string.IsNullOrEmpty(customCode)) //NestSweep(0,15,1,BinTogray)
                    {
                        sCode.SweepInfo = format;
                        sCode.Width = srclengthInPatInfo;
                        sCode.Order = Regex.Match(sweepCode, RegNestSweepCode, RegexOptions.IgnoreCase).Groups["order"].ToString().Trim();
                        return GenNestSweepRegisterAssign(sendBitName, sCode.Width, sCode.SweepInfo, sCode.Order);
                    }

                    var codeList = new List<string>();
                    foreach (string code in customCode.Split(','))
                    {
                        string tmpStr = code;
                        if (Regex.IsMatch(code, RegSpecialFormat, RegexOptions.IgnoreCase))
                        {
                            tmpStr = RegisterConvertor.ConvertNumberToSrc(sendBitName, code, patInfo);
                        }
                        codeList.Add(tmpStr);
                    }
                    sCode.SweepInfo = string.Join(",", codeList);
                    sCode.Order = Regex.Match(sweepCode, RegSweepCode1, RegexOptions.IgnoreCase).Groups["order"].ToString().Trim();
                    sCode.Algorithm = Regex.Match(sweepCode, RegSweepCode1, RegexOptions.IgnoreCase).Groups["alg"].ToString().Trim();
                    return GenRegisterAssign(sendBitName, sCode.SweepInfo, sCode.Order, sCode.Algorithm);
                }
                if (sendBitName.Equals("repeat"))
                {
                    sCode.SendBitName = string.Join(";", patInfo.SendBitName.Split('+').ToList());
                }


                sCode.Order = order;
                //sCode.Misc = Misc;
                sCode.Width = srclengthInPatInfo;

                if (!pattern.SweepCodes.ContainsKey(sCode.Order))
                {
                    pattern.SweepCodes.Add(sCode.Order, new List<SweepCode>());
                }

                pattern.SweepCodes[sCode.Order].Add(sCode);
            }
            string actualAssign = GenRegisterAssign(sendBitName, srclengthInPatInfo);
            return actualAssign;
        }

        internal string GenRegisterAssign(string sendBitName, int length)
        {
            if (length == 0)
            {
                return "";
            }

            string result = sendBitName + "=";
            for (int i = 0; i < length; i++)
            {
                result += "0";
            }
            return result;
        }

        internal string GenNestSweepRegisterAssign(string sendBitName, int width, string sweepInfo, string order)
        {
            //bit size@Start_Value@Stop_Value@Step_Value@ Algorithm(Option)
            return sendBitName + "=nestsweep" + order + "(" + width + "@" + sweepInfo + ")";
        }

        internal string GenRegisterAssign(string sendBitName, string sweepInfo, string order, string alg = "")
        {
            return string.IsNullOrEmpty(alg)
                ? sendBitName + "=sweep" + order + "[" + sweepInfo + "]"
                : sendBitName + "=sweep" + order + "[" + sweepInfo + "]@" + alg;
        }
    }
}
