using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Enums;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.EVS;

namespace TestPlanLib.Utility
{
    internal static partial class BinCutInstanceSheetReaderHelpers
    {
        private const string ConHeaderEvsNumConds = "_Num_Conds";
        private const string ConHeaderEvsPulses = "_Pulses";
        private const string ConHeaderEvsTime = "_Time";
        private const string ConHeaderEvsVoltage1 = "_Voltage1";
        private const string ConHeaderEvsVoltage2 = "_Voltage2";
        private const string ConHeaderEvsCooling = "_Cooling";
        private const string ConHeaderEvsCoolingAfter = "_Cooling_After";
        private const string ConHeaderEvsTotalPwr = "_TotalPwr";
        private const string ConHeaderEvsAlarmFlag = "_Alarm_Flag";
        private const string ConHeaderEvsRisingDelayTimeSec = "_RisingDelayTimeSec";
        private const string ConHeaderEvsRampUserFunction = "_Ramp_UserFunction";
        private const string ConHeaderEvsVtrigUserFunction = "_Vtrig_UserFunction";
        private const string ConHeaderEvsIVUserFunction = "_IV_UserFunction";

        [GeneratedRegex(@"(?<key>[a-zA-Z]+\d*)(?<codes>\[.+\])", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"(?<codes>\d+:\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"\d", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(@"\s+", RegexOptions.Compiled)]
        private static partial Regex MyRegex3();

        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex3 = MyRegex1();
        private static readonly Regex _regex4 = MyRegex2();
        private static readonly Regex _regex5 = MyRegex3();

        internal static string SplitByCode(string name, int colNumber, BinCutInstanceSheet binCutInstanceSheet, BinCutInstanceRow binCutInstanceRow, int index, int count)
        {
            if (_regex.IsMatch(name))
            {
                Match regexResult = _regex.Match(name);
                string codes = regexResult.Groups["codes"].ToString();
                List<string> iterates = GetCodes(codes);
                if (iterates.Count != count)
                {
                    binCutInstanceSheet.AddError(ScanErrorType.E_FormatError_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, colNumber, $"The code mismatch {name} !!!", [name]);
                    return name;
                }
                return name.Replace(codes, iterates[index]);
            }
            return name;
        }

        internal static List<string> GetCodes(string codesText)
        {
            var codes = new List<string>();
            string[] codeArray = codesText.TrimStart('[').TrimEnd(']').Split(',');
            foreach (string item in codeArray)
            {
                if (_regex3.IsMatch(item))
                {
                    string[] arr = item.Split(':');
                    if (int.TryParse(arr[0], out int start) &&
                        int.TryParse(arr.Last(), out int end))
                    {
                        for (int i = start; i <= end; i++)
                        {
                            codes.Add(i.ToString());
                        }
                    }
                }
                else if (_regex4.IsMatch(item))
                {
                    codes.Add(item);
                }
            }
            return codes;
        }

        internal static List<BinCutInstanceRow> MergebyInstanceName(List<BinCutInstanceRow> binCutInstanceRows)
        {
            var newList = new List<BinCutInstanceRow>();
            for (int i = 0; i < binCutInstanceRows.Count; i++)
            {
                BinCutInstanceRow row1 = binCutInstanceRows[i];
                if (row1.RowNum == 77)
                {
                }
                Dictionary<int, string> tempInitsWithIndex = row1.PatWithIndex;
                if (row1 == null)
                {
                    continue;
                }

                for (int j = i + 1; j < binCutInstanceRows.Count; j++)
                {
                    BinCutInstanceRow row2 = binCutInstanceRows[j];
                    if (row2.RowNum == 78)
                    {
                    }

                    if (!(row1.Instance.EqualsIgnoreCase(row2.Instance) &&
                          row1.FlowName.EqualsIgnoreCase(row2.FlowName)))
                    {
                        break;
                    }

                    if (BinCutInstanceRowUtility.IsConditionAllSame(row1, row2))
                    {
                        //Remove empty when TTR
                        row1.PatternList.AddRange(row2.PatternList.Where(x => !string.IsNullOrEmpty(x)));
                        row1.PayloadList.AddRange(row2.InitList.Where(x => !string.IsNullOrEmpty(x)));
                        row1.PayloadList.AddRange(row2.PayloadList.Where(x => !string.IsNullOrEmpty(x)));

                        //if (IsSameInit(row1, row2))
                        //    row1.PatternList.AddRange(row2.PayloadList);
                        //else
                        //{
                        //    row1.PatternList.AddRange(row2.PatternList);
                        //    row1.InitList.AddRange(row2.InitList);
                        //}
                        i = j;
                    }
                    else
                    {
                        i = j - 1;
                        break;
                    }
                }

                newList.Add(row1);
            }
            return newList;
        }

        internal static void FindTheFirstPayload(List<string> patListOrg, out int firstPayloadCnt, out bool existPayload)
        {
            #region find the first payload
            firstPayloadCnt = 0;
            existPayload = false;
            for (int idx = 0; idx < patListOrg.Count; idx++)
            {
                string pattern = patListOrg[idx];
                if (pattern == "NA")
                {
                    continue;
                }
                EnumPatternType type = pattern.GetPatternType();
                if (type == EnumPatternType.Payload)
                {
                    firstPayloadCnt = idx;
                    existPayload = true;
                    break;
                }
            }
            #endregion
        }

        internal static void FindTheLastInitPattern(ref int lastInitPatternCnt, ref bool existInit, List<string> patListOrg)
        {
            #region find the last init pattern
            for (int idx = 0; idx < patListOrg.Count; idx++)
            {
                string pattern = patListOrg[idx];
                if (string.IsNullOrEmpty(pattern) || pattern == "NA")
                {
                    continue;
                }

                EnumPatternType type = pattern.GetPatternType();
                if (type == EnumPatternType.Init)
                {
                    existInit = true;
                    lastInitPatternCnt = idx;
                }
            }
            #endregion
        }

        internal static void SetEvsConditionColIdx(int i, string text, string jobStage, string evsNum, List<EvsConditionColIdx> evsConditionColIdxs)
        {
            EvsConditionColIdx target = evsConditionColIdxs.First(x => x.JobStage == jobStage && x.EvsNum == evsNum);
            string prefix = jobStage + "_" + evsNum;
            if (text.EqualsIgnoreCase(prefix + ConHeaderEvsNumConds))
            {
                target.NumCondsIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsPulses))
            {
                target.PulsesIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsTime))
            {
                target.TimeIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsVoltage1))
            {
                target.Voltage1Idx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsVoltage2))
            {
                target.Voltage2Idx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsCooling))
            {
                target.CoolingIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsCoolingAfter))
            {
                target.CoolingAfterIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsTotalPwr))
            {
                target.TotalPwrIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsAlarmFlag))
            {
                target.AlarmFlagIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsRisingDelayTimeSec))
            {
                target.RisingDelayTimeSecIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsRampUserFunction))
            {
                target.RampIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsVtrigUserFunction))
            {
                target.VtrigIdx = i;
            }
            else if (text.EqualsIgnoreCase(prefix + ConHeaderEvsIVUserFunction))
            {
                target.IVdx = i;
            }
        }

        internal static void CheckPatternBase(ref BinCutInstanceSheet binCutInstanceSheet, ErrorCode emptyStringErrorType, ErrorCode firstPatternEmptyErrorType, ErrorCode patternEmptyErrorType)
        {
            foreach (BinCutInstanceRow row in binCutInstanceSheet.Rows)
            {
                bool hasEmpty = false;
                for (int i = 0; i < row.PatternList.Count; i++)
                {
                    if (_regex5.IsMatch(row.PatternList[i]))
                    {
                        if (!row.PatternList[i].Contains('+'))
                        {
                            binCutInstanceSheet.AddError(emptyStringErrorType, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.PatternStartColNumber + i, "An empty string exists in the pattern: " + row.PatternList[i] + " !!!", [row.PatternList[i]]);
                        }

                        hasEmpty = true;
                        row.PatternList[i] = _regex5.Replace(row.PatternList[i], "");
                    }

                    if (string.IsNullOrEmpty(row.PatternList[i]))
                    {
                        if (i == 0)
                        {
                            binCutInstanceSheet.AddError(firstPatternEmptyErrorType, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.PatternStartColNumber, "The fisrt pattern can not be empty !!!");
                        }
                        else
                        {
                            binCutInstanceSheet.AddError(patternEmptyErrorType, binCutInstanceSheet.SheetName, row.RowNum, binCutInstanceSheet.PatternStartColNumber + i, "The pattern can not be empty !!!");
                        }
                    }
                }

                if (hasEmpty)
                {
                    for (int i = 0; i < row.PayloadList.Count; i++)
                    {
                        row.PayloadList[i] = _regex5.Replace(row.PayloadList[i], "");
                    }
                }
            }
        }

        internal static int ClassifyPattern(EnumInstanceSheetType enumInstanceSheetType, int patternStartColNumber, BinCutInstanceSheet binCutInstanceSheet, int lastInitPatternCnt, bool existInit, BinCutInstanceRow binCutInstanceRow, List<string> patListOrg, int firstPayloadCnt, bool existPayload, int patternIdx)
        {
            for (int idx = 0; idx < patListOrg.Count; idx++)
            {
                string pattern = patListOrg[idx];
                if (pattern == "NA" || string.IsNullOrEmpty(pattern))
                {
                    continue;
                }

                EnumPatternType pattype = pattern.GetPatternType();
                if (pattern.Split('_').Length < 11 && !pattype.Equals(EnumPatternType.RetentionWait))
                {
                    if (enumInstanceSheetType == EnumInstanceSheetType.Bincut)
                    {
                        binCutInstanceSheet.AddError(BinCutErrorType.E_Pattern_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, idx + patternStartColNumber, $"The pattern : {pattern} format illegal", [pattern]);
                    }
                    else if (enumInstanceSheetType == EnumInstanceSheetType.Scan)
                    {
                        binCutInstanceSheet.AddError(ScanErrorType.E_Pattern_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, idx + patternStartColNumber, $"The pattern : {pattern} format illegal", [pattern]);
                    }
                    else if (enumInstanceSheetType == EnumInstanceSheetType.Evs)
                    {
                        binCutInstanceSheet.AddError(EvsErrorType.E_Pattern_01, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, idx + patternStartColNumber, $"The pattern : {pattern} format illegal", [pattern]);
                    }

                    continue;
                }
                binCutInstanceRow.Type = BincutInstanceType.Pattern;
                string[] spt = binCutInstanceRow.FlowName.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                bool isRtos = false;
                foreach (string s in spt)
                {
                    if (s == "RTOS")
                    {
                        isRtos = true;
                        break;
                    }
                }

                if (pattype == EnumPatternType.RetentionWait)
                {
                    binCutInstanceRow.RetentionWaitTime = pattern.Split('_').Last();
                    binCutInstanceRow.RetentionWaitIdx.Add(patternIdx);
                    patternIdx++;
                    continue;
                }

                if (pattype == EnumPatternType.Unknown && isRtos)
                {
                    binCutInstanceSheet.AddError(BinCutErrorType.E_Pattern_03, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, idx + patternStartColNumber, "The type of patten " + pattern + " is unknown !!!", [pattern]);
                    pattype = EnumPatternType.RTOS;
                }

                if (existPayload && idx > firstPayloadCnt)
                {
                    pattype = EnumPatternType.Payload;
                }

                if (pattype == EnumPatternType.Init)
                {
                    binCutInstanceRow.PatternList.Add(pattern);
                    binCutInstanceRow.InitList.Add(pattern);
                    binCutInstanceRow.PatWithIndex.Add(idx, pattern);
                }
                else if (pattype == EnumPatternType.Payload)
                {
                    binCutInstanceRow.PatternList.Add(pattern);
                    binCutInstanceRow.PayloadList.Add(pattern);
                    binCutInstanceRow.PatWithIndex.Add(idx, pattern);
                }
                else if (pattype == EnumPatternType.RTOS)
                {
                    binCutInstanceRow.Type = BincutInstanceType.Rtos;
                    binCutInstanceRow.PatternList.Add(pattern);
                    binCutInstanceRow.PayloadList.Add(pattern);
                    binCutInstanceRow.PatWithIndex.Add(idx, pattern);
                }
                else if (pattype == EnumPatternType.HARDIP)
                {
                    binCutInstanceRow.Type = BincutInstanceType.Hardip;
                    binCutInstanceRow.PatternList.Add(pattern);
                    binCutInstanceRow.PayloadList.Add(pattern);
                    binCutInstanceRow.PatWithIndex.Add(idx, pattern);
                }
                else
                {
                    if (existInit && lastInitPatternCnt < idx)
                    {
                        binCutInstanceRow.PatternList.Add(pattern);
                        binCutInstanceRow.PayloadList.Add(pattern);
                        binCutInstanceRow.PatWithIndex.Add(idx, pattern);
                    }
                    else
                    {
                        binCutInstanceRow.PatternList.Add(pattern);
                        binCutInstanceRow.InitList.Add(pattern);
                        binCutInstanceRow.PatWithIndex.Add(idx, pattern);
                    }

                    if (!string.IsNullOrEmpty(pattern))
                    {
                        binCutInstanceSheet.AddError(BinCutErrorType.E_Pattern_04, binCutInstanceSheet.SheetName, binCutInstanceRow.RowNum, idx + patternStartColNumber, "The type of patten " + pattern + " is unknown !!!", [pattern]);
                    }
                }
                patternIdx++;
            }

            return patternIdx;
        }
    }
}
