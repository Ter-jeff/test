using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.common.ReaderWriter.Reader.InputDataBase;

namespace DebugPlanReaderLib.DebugPlan.DigSrc
{
    public class DigSrcHandler
    {
        private List<HardIpReference> _patInfoData;
        private bool _isCSharp = false;
        public DigSrcHandler(List<HardIpReference> patInfoData, bool isCSharp = false)
        {
            _patInfoData = patInfoData;
            _isCSharp = isCSharp;
        }

        public void GenDigSrcIntoInstance(List<AiTestPlanSheet> planSheets)
        {
            //todo
            // 1. Entry point for setting DigSrc information for per plan item row
            // 2. Need to do ConvertCSharpPatternDigSrc after SetupDigSrcForPatCell

            foreach (AiTestPlanSheet sheet in planSheets)
            {
                foreach (AiTestPlanRow row in sheet.Rows)
                {
                    SetupSelsramForDigSrc(row);

                    if (sheet.IndexDigSrc != -1)
                    {
                        SetupDigSrcByPlanRow(row);
                    }
                }
            }
        }
        private void SetupSelsramForDigSrc(AiTestPlanRow row)
        {
            foreach (PatternDate pattern in row.Patterns)
            {
                if (pattern.Name.ToUpper().Contains("_SRMDSSC"))
                {
                    HardIpReference targetPatInfo = null;
                    if (string.IsNullOrEmpty(pattern.Version))
                    {
                        targetPatInfo = _patInfoData.LastOrDefault(x => x.Payload.ToUpper() == pattern.Name.ToUpper());
                    }
                    else
                    {
                        targetPatInfo = _patInfoData.FirstOrDefault(x => x.Payload.ToUpper() == pattern.Name.ToUpper() && x.Version.ToUpper() == pattern.Version.ToUpper());
                    }
                    if (targetPatInfo == null)
                    {
                        //SetupDigSrcSRMPlanforNewTChar(pattern, row);
                        continue;
                    }

                    if (string.IsNullOrEmpty(targetPatInfo.SendBitStr))
                    {
                        //SetupDigSrcSRMPlanforNewTChar(pattern, row);
                        continue;
                    }

                    if (!_isCSharp)
                    {
                        pattern.DigSrcBitSize = targetPatInfo.SendBit.ToString();
                        //pattern.DigSrcEQ = targetPatInfo.SendBitStr;
                        pattern.DigSrcEQ = "sgmt0_" + pattern.DigSrcBitSize;
                        pattern.DigSrcPin = targetPatInfo.SendPinName;
                        pattern.DigSrcSeg = "sgmt0=SELSRAM";
                        pattern.SelsramDigSrc = true;

                    }
                    else
                    {
                        //Gen for CSharp version
                        string selsrmBits = Regex.Replace(row.SelsramDssc, @"selsrm", "", RegexOptions.IgnoreCase);
                        string selsrmBitsReplace = Regex.Replace(selsrmBits, @"s", "", RegexOptions.IgnoreCase);
                        string sgmt = "";
                        if (!string.IsNullOrEmpty(selsrmBitsReplace))
                        {
                            sgmt = $"sgmt0_=SELSRAM({string.Join(",", selsrmBits.ToArray())})";
                        }
                        else
                        {
                            sgmt = $"sgmt0=SELSRAM()";
                        }

                        pattern.DigSrcBitSize = targetPatInfo.SendBit.ToString();
                        //pattern.DigSrcEQ = targetPatInfo.SendBitStr;
                        pattern.DigSrcEQ = "sgmt0_" + pattern.DigSrcBitSize;
                        pattern.DigSrcPin = targetPatInfo.SendPinName;
                        pattern.DigSrcSeg = sgmt;
                        pattern.SelsramDigSrc = true;

                    }
                }
            }
        }

        private void SetupDigSrcSRMPlanforNewTChar(PatternDate patternDate, AiTestPlanRow row)
        {
            if (!string.IsNullOrEmpty(row.SelsramDssc.Trim()))
            {
                var selsramSplit = Regex.Split(row.SelsramDssc.Trim(), @"Selsr[a]*m", RegexOptions.IgnoreCase);
                if (selsramSplit.Length > 1)
                {
                    var selsramStrLength = selsramSplit[1].Length;
                    patternDate.DigSrcBitSize = selsramStrLength.ToString();
                    patternDate.DigSrcEQ = "sgmt0_" + patternDate.DigSrcBitSize;
                    patternDate.DigSrcPin = "JTAG_TDI";
                    patternDate.DigSrcSeg = "sgmt0=SELSRAM";
                    patternDate.SelsramDigSrc = true;
                }
            }
        }

        private void SetupDigSrcByPlanRow(AiTestPlanRow row)
        {
            if (string.IsNullOrEmpty(row.DigSrc.Trim()))
            {
                return;
            }
            string[] splitByPat = row.DigSrc.Split(',');
            foreach (string split in splitByPat)
            {
                string[] settings = split.Split(':');
                if (settings.Length < 2)
                {
                    continue;
                }

                Match patSetting = Regex.Match(settings[0].Trim(), @"^(Pat|Pattern)(?<index>\d)$+", RegexOptions.IgnoreCase);
                string bitSetting = settings[1].Trim();
                if (!patSetting.Success)
                {
                    continue;
                }

                if (string.IsNullOrEmpty(bitSetting))
                {
                    continue;
                }

                string patIndex = patSetting.Groups["index"].Value;
                PatternDate targetPat = row.Patterns.FirstOrDefault(x => x.Index == patIndex);
                if (targetPat == null)
                {
                    continue;
                }

                HardIpReference targetPatInfo = null;
                if (string.IsNullOrEmpty(targetPat.Version))
                {
                    targetPatInfo = _patInfoData.LastOrDefault(x => x.Payload.ToUpper() == targetPat.Name.ToUpper());
                }
                else
                {
                    targetPatInfo = _patInfoData.FirstOrDefault(x => x.Payload.ToUpper() == targetPat.Name.ToUpper() && x.Version.ToUpper() == targetPat.Version.ToUpper());
                }

                if (targetPatInfo == null)
                {
                    continue;
                }

                targetPat.SelsramDigSrc = false;
                SetupDigSrcForPatCell(bitSetting, targetPat, targetPatInfo);
            }
        }

        private static void ConvertCSharpPatternDigSrc(AiTestPlanRow planItem, Dictionary<string, string> ArgPatternIndexConvertedDsscDict)
        {
            //digSrcPin
            //digSrcAssignment
            //digSrcEquation

            var digSrcEquation = "";
            var digSrcAssignment = new List<string>();
            var digSrcPin = "";

            string[] patternIndexList = new string[planItem.Patterns.Count * 2 - 1];
            if (string.IsNullOrEmpty(planItem.DigSrcPin))
                return;
            var sourcePin = planItem.DigSrcPin.Split(',').ToList().Select(p => p.Split(':')[1]).ToList().Distinct().FirstOrDefault().ToString();
            var sourceEqList = planItem.DigSrcEQ.Split(',');
            var sourceSegList = planItem.DigSrcSeg.Split(',');
            //var sourceAssign = planItem.DigSrcAssignment.Split(',');

            for (int i = 0; i < patternIndexList.Length; ++i)
                patternIndexList[i] = i % 2 == 1 ? "|" : "";

            int sgmtIdx = 0;


            foreach (var _sourceSeg in sourceEqList)
            {
                var splitSegBit = ArgPatternIndexConvertedDsscDict[_sourceSeg.Split(':')[0]].Split(':')[0].Replace("INIT", "").Replace("PL", "");
                int bitIdx = -1;
                int.TryParse(splitSegBit, out bitIdx);

                patternIndexList[bitIdx * 2 - 2] = $"sgmt{sgmtIdx}";

                //if(sourceSegList[sgmtIdx].Split(';').Count() > 1)
                //{
                //    var binStr = ConvertBinStr("sgmt13f1sgmt14f1sgmt54f01000000sgmt58f01000000", patInfo, planItem.IsDateNeedReverse);

                //}

                var sgmt = sourceSegList[sgmtIdx].Split('=')[1];

                if (Regex.IsMatch(sgmt, @"SELSRAM", RegexOptions.IgnoreCase))
                {
                    var userdefLast = planItem.TestInstanceName.Split('_').Where(x => !string.IsNullOrEmpty(x)).ToList();

                    if (userdefLast.Count() > 10)
                    {
                        //vbtFunction.SetParamValue("selsrmDSSC", "\'" + userdefLast[10]);

                        var selsrmBits = Regex.Replace(userdefLast[10], @"selsrm", "", RegexOptions.IgnoreCase);
                        var selsrmBitsReplace = Regex.Replace(selsrmBits, @"s", "", RegexOptions.IgnoreCase);

                        var segment = planItem.DigSrcSeg.Split(':')[1].Split('=')[0];

                        if (!string.IsNullOrEmpty(selsrmBitsReplace))
                        {
                            sgmt = $"sgmt{sgmtIdx}=SELSRAM({string.Join(",", selsrmBits.ToArray())})";
                        }
                        else
                        {
                            sgmt = $"sgmt{sgmtIdx}=SELSRAM()";
                        }
                        digSrcAssignment.Add(sgmt);
                    }

                }
                else
                {
                    //sgmt = sourceAssign[bitIdx + 1];

                    if (string.IsNullOrEmpty(sgmt))
                    {
                        //sgmt = ConvertBinStr(sourceAssign[bitIdx - 1], patInfo, false);

                    }

                    digSrcAssignment.Add($"sgmt{sgmtIdx}={sgmt}");

                }
                sgmtIdx += 1;
            }
            digSrcEquation = string.Join("", patternIndexList);

            //vbtFunction.SetParamValue("digSrcPin", sourcePin);
            //vbtFunction.SetParamValue("digSrcEquation", digSrcEquation);
            //vbtFunction.SetParamValue("digSrcAssignment", string.Join(";", digSrcAssignment));
        }

        private void SetupDigSrcForPatCell(string bitSetting, PatternDate pat, HardIpReference patternInfo)
        {
            pat.DigSrcBitSize = patternInfo.SendBit.ToString();
            pat.DigSrcEQ = patternInfo.SendBitStr;
            pat.DigSrcPin = patternInfo.SendPinName;

            if (Regex.IsMatch(bitSetting, @"^[0|1]+$"))
            {
                SetDigSrcSegByBits(pat, bitSetting, patternInfo);
            }
            else
            {
                SetDigSrcSegBySgmts(pat, bitSetting, patternInfo);
            }
        }

        private void SetDigSrcSegByBits(PatternDate pat, string bits, HardIpReference patternInfo)
        {
            List<string> sgmtList = patternInfo.SendBitStr.Split('+').ToList();
            string unsetBit = bits;
            var resultList = new List<string>();
            foreach (string sgmt in sgmtList)
            {
                Match sgmtRegex = Regex.Match(sgmt, @"(?<header>sgmt\d+)_(?<length>\d+)", RegexOptions.IgnoreCase);
                if (!sgmtRegex.Success)
                {
                    continue;
                }

                string header = sgmtRegex.Groups["header"].Value;
                int length = int.Parse(sgmtRegex.Groups["length"].Value);
                if (unsetBit.Length < length)
                {
                    if (unsetBit.Length == 0)
                    {
                        break;
                    }
                    resultList.Add(string.Format($"{header}=0b{unsetBit.PadLeft(length, '0')}"));
                    break;
                }
                else
                {
                    resultList.Add(string.Format($"{header}=0b{unsetBit.Substring(0, length)}"));
                    unsetBit = unsetBit.Substring(length);
                    continue;
                }
            }

            pat.DigSrcSeg = string.Join(";", resultList);
            if (patternInfo.SendBit > bits.Length)
            {
                pat.DigSrcBits = bits.PadLeft(patternInfo.SendBit, '0');
            }
            else
            {
                pat.DigSrcBits = bits.Substring(bits.Length - patternInfo.SendBit, patternInfo.SendBit);
            }
        }

        private void SetDigSrcSegBySgmts(PatternDate pat, string sgmtStr, HardIpReference patternInfo)
        {
            List<string> sgmtList = patternInfo.SendBitStr.Split('+').ToList();
            List<string> sgmtSettingList = sgmtStr.Split(';').ToList();
            var resultList = new List<string>();
            var bitsDict = sgmtList.ToDictionary(x => x, x => "");

            foreach (string sgmtSetting in sgmtSettingList)
            {
                Match settingRegex = Regex.Match(sgmtSetting, @"^(?<header>sgmt\d+)=0b(?<bits>[0|1]+)$");

                if (!settingRegex.Success)
                {
                    continue;
                }

                string header = settingRegex.Groups["header"].Value;
                string bits = settingRegex.Groups["bits"].Value;

                string patInfoSgmt = sgmtList.FirstOrDefault(x => x.StartsWith(header + "_", StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrEmpty(patInfoSgmt))
                {
                    continue;
                }

                Match sgmtRegex = Regex.Match(patInfoSgmt, @"(?<header>sgmt\d+)_(?<length>\d+)", RegexOptions.IgnoreCase);

                if (!sgmtRegex.Success)
                {
                    continue;
                }

                int length = int.Parse(sgmtRegex.Groups["length"].Value);

                string result = "";
                if (length > bits.Length)
                {
                    result = bits.PadLeft(length, '0');
                }
                else if (length < bits.Length)
                {
                    result = bits.Substring(bits.Length - length, length);
                }
                else
                {
                    result = bits;
                }

                resultList.Add(string.Format($"{header}=0b{result}"));
                bitsDict[header + "_" + length] = result;
            }
            pat.DigSrcSeg = string.Join(";", resultList);

            foreach (string sgmt in sgmtList)
            {
                Match sgmtRegex = Regex.Match(sgmt, @"(?<header>sgmt\d+)_(?<length>\d+)", RegexOptions.IgnoreCase);

                if (!sgmtRegex.Success)
                {
                    continue;
                }

                string header = sgmtRegex.Groups["header"].Value;
                int length = int.Parse(sgmtRegex.Groups["length"].Value);
                if (bitsDict.ContainsKey(sgmt))
                {
                    if (string.IsNullOrEmpty(bitsDict[sgmt]))
                    {
                        bitsDict[sgmt] = "".PadLeft(length, '0');
                    }
                }
            }
            pat.DigSrcBits = string.Join("", bitsDict.Values);
        }
    }
}
