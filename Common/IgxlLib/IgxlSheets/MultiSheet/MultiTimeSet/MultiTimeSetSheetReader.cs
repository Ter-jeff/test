using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.Regs;
using IgxlLib.Utility;

namespace IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet
{
    public partial class MultiTimeSetSheetReader
    {
        public const string CycleS = "Cycle_S";
        public const string ClockS = "Clock_S";
        public const string ClockE = "Clock_E";
        public const string Strobe = "Strobe";

        private const string TimeModePattern = @"Timing Mode:[\t](?<str>\w*)";
        private const string MasterTsPattern = @"Master Timeset Name:[\t](?<str>\w*)";
        private const string TimeDomainPattern = @"Time Domain:[\t](?<str>\w*)";
        private const string Header = @"Time Set[\t]Period";

        [GeneratedRegex(@"_(?<var>[\d|\w]+)", RegexOptions.Compiled)]
        private static partial Regex ContextVariableTokenRegex();
        [GeneratedRegex(TimeModePattern, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex TimingModeRegex();
        [GeneratedRegex(MasterTsPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MasterTimeSetRegex();
        [GeneratedRegex(TimeDomainPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex TimeDomainRegex();
        [GeneratedRegex(".txt", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex TxtExtensionRegex();
        [GeneratedRegex(Header, RegexOptions.Compiled)]
        private static partial Regex HeaderRegex();
        [GeneratedRegex("DTTimesetBasicSheet,version=2.3", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex Version2P3Regex();

        private static readonly Regex _contextVariableTokenRegex = ContextVariableTokenRegex();
        private static readonly Regex _timingModeRegex = TimingModeRegex();
        private static readonly Regex _masterTimeSetRegex = MasterTimeSetRegex();
        private static readonly Regex _timeDomainRegex = TimeDomainRegex();
        private static readonly Regex _txtExtensionRegex = TxtExtensionRegex();
        private static readonly Regex _headerRegex = HeaderRegex();
        private static readonly Regex _version2P3Regex = Version2P3Regex();

        public static MultiTimeSetSheets? ReadTimeSetTxt1P4(List<string> timeSetPathList, bool isRemoveBackup = false)
        {
            var comTimeSetSheetCollection = new MultiTimeSetSheets();

            try
            {
                foreach (string timeSetPath in timeSetPathList)
                {
                    if (!File.Exists(timeSetPath))
                    {
                        continue;
                    }

                    string[] lines = File.ReadAllLines(timeSetPath);
                    string strobe = "";

                    string timeMode = "";
                    string masterTs = "";
                    string timeDomain = "";

                    if (lines.Length > 2)
                    {
                        timeMode = _timingModeRegex.Match(lines[2]).Groups["str"].ToString();
                        masterTs = _masterTimeSetRegex.Match(lines[2]).Groups["str"].ToString();
                    }
                    if (lines.Length > 4)
                    {
                        timeDomain = _timeDomainRegex.Match(lines[3]).Groups["str"].ToString();
                    }

                    TimeRow1P4Converter timeRowConverter = Converter(lines[0]);
                    string sheetName = _txtExtensionRegex.Replace(Path.GetFileName(timeSetPath), "");
                    var timeSetBasicSheet = new ComTimeSetBasicSheet(sheetName, timeMode, masterTs, timeDomain, strobe);
                    var pairs = new Dictionary<string, ComTimeSetBasic>();
                    int lIStartRowNum = 4;
                    for (int i = lIStartRowNum; i < lines.Length; i++)
                    {
                        if (_headerRegex.IsMatch(lines[i]))
                        {
                            lIStartRowNum = i + 1;
                            break;
                        }
                    }

                    ReadData(lines, timeRowConverter, sheetName, timeSetBasicSheet, pairs, lIStartRowNum);

                    CheckMissingEquationBaseVar(sheetName, timeSetBasicSheet, pairs);

                    foreach (KeyValuePair<string, ComTimeSetBasic> keyValuePair in pairs)
                    {
                        if (string.IsNullOrEmpty(keyValuePair.Value.Name) && isRemoveBackup)
                        {
                            break;
                        }

                        if (!string.IsNullOrEmpty(keyValuePair.Value.Name))
                        {
                            timeSetBasicSheet.AddRow(keyValuePair.Value);
                        }
                    }

                    comTimeSetSheetCollection.AddTimeSetSheet(timeSetBasicSheet);
                }
            }
            catch (Exception)
            {
                return null;
            }
            return comTimeSetSheetCollection;
        }

        internal static void ReadData(string[] lines, TimeRow1P4Converter timeRow1P4Converter, string sheetName, ComTimeSetBasicSheet comTimeSetBasicSheet, Dictionary<string, ComTimeSetBasic> pairs, int startRowNum)
        {
            bool startVarDefinitions = false;
            for (int i = startRowNum; i < lines.Length; i++)
            {
                string[] arr = lines[i].Split('\t');
                if (arr.Length == 0 || lines[i].Split(['\t'], StringSplitOptions.RemoveEmptyEntries).Length == 0)
                {
                    continue;
                }

                if (lines[i].Contains("VAR Definitions") && !startVarDefinitions)
                {
                    startVarDefinitions = true;
                    continue;
                }
                if (startVarDefinitions)
                {
                    //	HTOL_Freq_VAR	=1000000, tset1Per	=1000.000*ns
                    if (lines[i].Contains('='))
                    {
                        string[] spt = lines[i].Split('=');
                        string varTok = spt[0].Trim();
                        varTok = varTok.ReplaceStartsWith(CommonConst.UnderScore, "");
                        string valueTok = spt[1].Trim();
                        string convertUnit = valueTok.ConvertNumber();
                        //Use HardIp Function
                        bool isNumOk = double.TryParse(convertUnit, out double dValue);
                        if (isNumOk)
                        {
                            foreach (KeyValuePair<string, ComTimeSetBasic> subTsb in pairs)
                            {
                                if (subTsb.Value.SubContextVariable.Contains(varTok))
                                //which means the variable appear on the comment is use by this time set
                                {
                                    subTsb.Value.SubCommentVariable.TryAdd(varTok, dValue);
                                }
                            }
                        }
                        //Add error report...  under construct
                    }
                    else
                    {
                        string varTok = Reg.WhitespaceRegex.Replace(lines[i], "");
                        double dValue = -1e9;
                        foreach (KeyValuePair<string, ComTimeSetBasic> subTsb in pairs)
                        {
                            if (subTsb.Value.SubContextVariable.Contains(varTok))
                            //which means the variable appear on the comment is use by this time set
                            {
                                subTsb.Value.SubCommentVariable.TryAdd(varTok, dValue);
                            }
                        }
                    }

                    continue;
                }

                string timeSet = arr[1];
                string clockPeriod = arr[2];
                //ReadTimeRow1P4() add argument _contextVar for read equation base variable
                if (!ReadTimeRow(arr, timeRow1P4Converter, out TimingRow timingRow, out List<string> contextVar, out Dictionary<string, double> shiftFreqVar))
                {
                    comTimeSetBasicSheet.AddError(BasicErrorType.E_InvalidTiming_01, sheetName, i + 1, 0, [sheetName, (i + 1).ToString()]);
                    continue;
                }

                if (pairs.TryGetValue(timeSet, out ComTimeSetBasic? value))
                {
                    value.AddTimingRow(timingRow);
                    foreach (string varTmp in contextVar)
                    {
                        if (!pairs[timeSet].SubContextVariable.Contains(varTmp))
                        {
                            pairs[timeSet].SubContextVariable.Add(varTmp);
                        }
                    }

                    foreach (KeyValuePair<string, double> dicPair in shiftFreqVar)
                    {
                        if (!pairs[timeSet].ShiftInReserve.ContainsKey(dicPair.Key))
                        {
                            pairs[timeSet].ShiftInReserve.Add(dicPair.Key, dicPair.Value);
                        }
                    }
                }
                else
                {
                    var timeSetBasic = new ComTimeSetBasic { Name = timeSet, CyclePeriod = clockPeriod };
                    timeSetBasic.AddTimingRow(timingRow);

                    foreach (string varTmp in contextVar)
                    {
                        timeSetBasic.SubContextVariable.Add(varTmp);
                    }

                    foreach (KeyValuePair<string, double> dicPair in shiftFreqVar)
                    {
                        timeSetBasic.ShiftInReserve.Add(dicPair.Key, dicPair.Value);
                    }

                    pairs.Add(timeSet, timeSetBasic);
                }
            }
        }

        internal static void CheckMissingEquationBaseVar(string sheetName, ComTimeSetBasicSheet comTimeSetBasicSheet, Dictionary<string, ComTimeSetBasic> dic)
        {
            foreach (KeyValuePair<string, ComTimeSetBasic> tSetDataPair in dic)
            {
                List<string> contextVars = tSetDataPair.Value.SubContextVariable;
                Dictionary<string, double> commentVarsDict = tSetDataPair.Value.SubCommentVariable;

                foreach (KeyValuePair<string, double> commentPair in commentVarsDict)
                {
                    if (Math.Abs(commentPair.Value - -1e9) < CommonConst.Tolerance)
                    {
                        //string errMsg = $"Equation base variable '{commentPair.Key}' used in Time Set file {sheetName} is not assigned an initial value";
                        //var error = new Error(EnumErrorType.E_FormatError_04, ErrorLevel.Error, sheetName, 0, 0, $"Equation base variable '{commentPair.Key}' used in Time Set file {sheetName} is not assigned an initial value");
                        comTimeSetBasicSheet.AddError(BasicErrorType.E_InvalidTiming_02, sheetName, 0, 0, [commentPair.Key, sheetName]);
                    }
                }

                //Check Rule2. use context vars as base, check if comment vars are not equal to context
                var commentVars = commentVarsDict.Keys.ToList();
                foreach (string contextVar in contextVars)
                {
                    if (!commentVars.Contains(contextVar))
                    {
                        //string errMsg = $"Equation base variable '{contextVar}' used in the context of Time Set file {sheetName} is not assigned value in comment";
                        //var error = new Error(EnumErrorType.E_FormatError_05, ErrorLevel.Error, sheetName, 0, 0, $"Equation base variable '{contextVar}' used in the context of Time Set file {sheetName} is not assigned value in comment");
                        comTimeSetBasicSheet.AddError(BasicErrorType.E_InvalidTiming_03, sheetName, 0, 0, [contextVar, sheetName]);
                    }
                }
            }
        }

        internal static bool ReadTimeRow(string[] line, TimeRow1P4Converter timeRow1P4Converter, out TimingRow timingRow, out List<string> subContextVar, out Dictionary<string, double> shiftInReserveVar)
        {
            subContextVar = [];
            shiftInReserveVar = [];

            timingRow = timeRow1P4Converter.ConvertTimeRow(line);

            GetContextVariable(timingRow.PinGrpClockPeriod, ref subContextVar);
            GetContextVariable(timingRow.DriveOn, ref subContextVar);
            GetContextVariable(timingRow.DriveData, ref subContextVar);
            GetContextVariable(timingRow.DriveReturn, ref subContextVar);
            GetContextVariable(timingRow.DriveOff, ref subContextVar);
            GetContextVariable(timingRow.CompareOpen, ref subContextVar);
            GetContextVariable(timingRow.CompareClose, ref subContextVar);

            try
            {
                string cyclePeriod = line[2];
                if (!string.IsNullOrEmpty(cyclePeriod))
                {
                    if (IsContextVariable(cyclePeriod))
                    {
                        GetContextVariable(cyclePeriod, ref subContextVar);
                    }
                    else
                    {
                        decimal periodVal = GetEgValueInDecimal(cyclePeriod);
                        if (periodVal != (decimal)0.0 && !string.IsNullOrEmpty(timingRow.DriveData))
                        {

                            decimal d1Val = GetEgValueInDecimal(timingRow.DriveData);
                            decimal tRatio = d1Val / periodVal;
                            tRatio = Math.Round(tRatio, 2);
                            if (timingRow.DataFmt.ContainsIgnoreCase("RL"))
                            {
                                if (!shiftInReserveVar.ContainsKey(SpecFormat.GenAcSpecSymbol(CycleS)))
                                {
                                    shiftInReserveVar.Add(SpecFormat.GenAcSpecSymbol(CycleS), (double)tRatio);
                                }
                            }
                        }

                        if (periodVal != (decimal)0.0 && !string.IsNullOrEmpty(timingRow.CompareOpen))
                        {
                            decimal c0Val = GetEgValueInDecimal(timingRow.CompareOpen);
                            decimal tRatio = c0Val / periodVal;
                            tRatio = Math.Round(tRatio, 2);

                            if (tRatio != 0)
                            {
                                if (!shiftInReserveVar.ContainsKey(SpecFormat.GenAcSpecSymbol(Strobe)))
                                {
                                    shiftInReserveVar.Add(SpecFormat.GenAcSpecSymbol(Strobe), (double)tRatio);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                //throw new Exception(string.Join("", line));
                return false;
            }

            return !timeRow1P4Converter.NeedCompensate(line);
        }

        private static decimal GetEgValueInDecimal(string inStr)
        {
            decimal dVal = inStr.Contains('E') ? Convert.ToDecimal(decimal.Parse(inStr, NumberStyles.Float)) : decimal.Parse(inStr);
            return dVal;
        }

        private static void GetContextVariable(string cell, ref List<string> subContextVar)
        {
            //cell context example:
            //=_RT_CLK32768_Freq_GLB 
            //=(1/_TCK_Freq_VAR)
            //=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR
            //=_Cycle_S_VAR+0.7/_ShiftIn_Freq_VAR+_Clock_E_VAR

            MatchCollection matches = _contextVariableTokenRegex.Matches(cell);
            foreach (Match match in matches)
            {
                string contextVar = match.Groups["var"].ToString();
                if (!string.IsNullOrEmpty(contextVar) && !subContextVar.Contains(contextVar))
                {
                    subContextVar.Add(contextVar);
                }
            }
        }

        private static bool IsContextVariable(string cell)
        {
            //cell context example:
            //=_RT_CLK32768_Freq_GLB 
            //=(1/_TCK_Freq_VAR)
            //=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR
            //=_Cycle_S_VAR+0.7/_ShiftIn_Freq_VAR+_Clock_E_VAR    
            return _contextVariableTokenRegex.IsMatch(cell);
        }

        public static TimeRow1P4Converter Converter(string header)
        {
            if (_version2P3Regex.IsMatch(header))
            {
                return new TimeRow2P3Converter();
            }
            return new TimeRow1P4Converter();
        }
    }
}
