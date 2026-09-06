using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using EfuseCheckCmdLib.IgxlLogLib.Base;

namespace EfuseCheckCmdLib.IgxlLogLib
{
    public partial class TestDataFormat
    {

        private readonly Regex _regexFail = FailRegex();
        private readonly Regex _regexAlarm = AlarmRegex();

        [GeneratedRegex(@"\(F\)")]
        private static partial Regex FailRegex();

        [GeneratedRegex(@"\(A\)")]
        private static partial Regex AlarmRegex();

        [GeneratedRegex(@"(?<Title>.*Number\s+Site.*1st Failed Cycle.*)")]
        private static partial Regex TitleFunctionalRegex();

        [GeneratedRegex(@"(?<Title>.*Number\s+Site.*Measured.*)")]
        private static partial Regex TitleMeasurementRegex();

        [GeneratedRegex(",")]
        private static partial Regex CommaRegex();

        [GeneratedRegex(@"\.pat|\.gz", RegexOptions.IgnoreCase)]
        private static partial Regex PatOrGzRegex();

        [GeneratedRegex(@".+\\")]
        private static partial Regex BackslashPrefixRegex();

        [GeneratedRegex(@"\.pat\.gz|\.pat", RegexOptions.IgnoreCase)]
        private static partial Regex PatGzSuffixRegex();

        //記錄所有 Data format with different index length
        private readonly Dictionary<int, Dictionary<string, DataLogFormatHeader>> _lineLengthTestNumFormatFunctionDict = [];

        private readonly HashSet<int> _functionHeaderCnt = [];

        private readonly Dictionary<int, Dictionary<string, DataLogFormatHeader>> _lineLengthTestNumFormatParametricDict = [];

        private readonly HashSet<int> _parametricHeaderCnt = [];
        //知道現在的Test Num Print是哪一種格式 輸入Test Name知要截斷的開頭結尾

        private Dictionary<string, DataLogFormatHeader>? _currentTestNumFormat = [];

        public bool IsFormatLine(string line)
        {
            if (line.IndexOf("Number", StringComparison.Ordinal) > -1 && line.IndexOf("Site", StringComparison.Ordinal) > -1)
            {
                // currTd.CurrentTnFormatMeas;
                _currentTestNumFormat = BulidTestNumFormat(line);
                if (_currentTestNumFormat != null)
                {
                    // record all length format
                    if (_currentTestNumFormat.ContainsKey("Pattern"))
                    {
                        _lineLengthTestNumFormatFunctionDict.TryAdd(_currentTestNumFormat["Max Length"].Length, _currentTestNumFormat);
                    }
                    if (_currentTestNumFormat.ContainsKey("Measured") || _currentTestNumFormat.ContainsKey("Pin"))
                    {
                        _lineLengthTestNumFormatParametricDict.TryAdd(_currentTestNumFormat["Max Length"].Length, _currentTestNumFormat);
                    }
                }
                return true;
            }
            return false;
        }

        private Dictionary<string, DataLogFormatHeader>? BulidTestNumFormat(string titalLine) //當遇到Number 開頭的Line 就知道接下來的Test Numeber Format!!
        {
            //Key = Item , DataLogFormatHeader.StartIdx = 起始點, DataLogFormatHeader.Length = 長度 給標準SubString參數用的
            var dicItemStartLength = new Dictionary<string, DataLogFormatHeader>();

            EDataFormatType headerType;

            string currTitle;
            if (titalLine.IndexOf("Pattern", StringComparison.Ordinal) > -1)
            {
                currTitle =
                    TitleFunctionalRegex().Match(titalLine).Groups["Title"].ToString();
                headerType = EDataFormatType.Functional;
            }

            else if (titalLine.IndexOf("Measured", StringComparison.Ordinal) > -1)
            {
                currTitle = TitleMeasurementRegex().Match(titalLine).Groups["Title"].ToString();
                headerType = EDataFormatType.Measurment;
            }
            else
            {
                return null;
            }

            var possibleItems = new List<string>() { "Number", "Site", "Test Name", "Pin", "Channel", "Low", "Measured", "High",
                "Force", "Loc", "Pattern", "1st Failed Cycle", "Total Failed Cycles" };

            //var activeTitle = 0; //用來統計應該有多少個欄位 用來判斷列印出來應該有多少值

            var startIndex = new List<int>();
            foreach (string item in possibleItems) //第一次先找到各個項目的起始點
            {
                //先找到起始點//+ 1
                int sIndex = currTitle.IndexOf(item, StringComparison.Ordinal);
                if (sIndex != -1) //-1代表根本沒有這個Item
                {
                    dicItemStartLength[item] = new DataLogFormatHeader { HeaderName = item, StartIdx = sIndex };
                    startIndex.Add(sIndex);
                }
            }
            //這是給最後的Item定位用的

            startIndex.Add(currTitle.Length);
            //起始點 由小到大排序
            startIndex.Sort();

            foreach (string item in dicItemStartLength.Keys)
            {
                int len = currTitle.Length;
                foreach (int i in startIndex)
                {
                    if (dicItemStartLength[item].StartIdx < i)
                    {
                        if (i - dicItemStartLength[item].StartIdx <= len)
                        {
                            //相減最小的就是下一個起始點 
                            len = i - dicItemStartLength[item].StartIdx;
                        }
                    }
                }
                dicItemStartLength[item].Length = len - 1;
            }
            //var splitCnt = Regex.Split(titalLine, @"\s+").Count();

            switch (headerType)
            {
                case EDataFormatType.Measurment:
                    _parametricHeaderCnt.Add(dicItemStartLength.Count);
                    break;
                case EDataFormatType.Functional:
                    _functionHeaderCnt.Add(dicItemStartLength.Count);
                    break;
            }

            dicItemStartLength["Max Length"] = new DataLogFormatHeader { StartIdx = currTitle.Length, Length = currTitle.Length, HeaderName = "Max Length" };

            return dicItemStartLength;
        }

        private static bool IsShmooLine(string line)
        {
            if (line.IndexOf("[Char", StringComparison.Ordinal) > -1 && line.IndexOf(']') > -1 &&
                line.IndexOfIgnoreCase("pat") > -1)
            {
                return true;
            }

            return false;

        }

        public DataFormatDataRow? GetDataRow(string line, string? currentInstance, bool functionOnly, int actuallineNum = 0)
        {
            int splitCnt = line.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Length;
            int functionHeaderMaxCnt = _functionHeaderCnt.Count != 0 ? _functionHeaderCnt.ToList().Max() : 0;
            int parametricHeaderMaxCnt = _parametricHeaderCnt.Count != 0 ? _parametricHeaderCnt.ToList().Min() : 0;
            Dictionary<string, DataLogFormatHeader>? cfForamtDic = GetCfForamtDic(line, splitCnt, functionHeaderMaxCnt, parametricHeaderMaxCnt);

            if (cfForamtDic == null)
            {
                return null;
            }

            try
            {
                int activeSite = 999;
                string testname = "";
                string pin = "";
                string channel = "";
                string low = "";
                string measured = "";
                string high = "";
                string force = "";
                string loc = "";
                string lowUnit = "";
                string highUnit = "";
                string measUnit = "";
                string forceUnit = "";
                long number = 000;
                string unit = "";
                EDataFormatType dataFormat;
                string pat = "";
                string testInstance = currentInstance ?? "";
                string testresult = _regexFail.IsMatch(line) || _regexAlarm.IsMatch(line) ? "Fail" : "Pass";

                #region Measuremnt Type
                if (IsShmooLine(line))
                {
                    if (functionOnly)
                    {
                        return null;
                    }

                    activeSite = GetShmooLine(line, activeSite, out testname, out number, out dataFormat, out pat, out testInstance);
                }
                else if (cfForamtDic.ContainsKey("Measured")) //Measuremnt Type
                {
                    if (functionOnly)
                    {
                        return null;
                    }

                    dataFormat = EDataFormatType.Measurment;
                    if (cfForamtDic["Max Length"].Length == line.Length || cfForamtDic["Max Length"].Length == line.Length + 1)
                    {
                        #region The same format as header
                        string tmpString = line.Substring(cfForamtDic["Number"].StartIdx, cfForamtDic["Number"].Length).Trim();
                        if (!long.TryParse(tmpString, out number))
                        {
                            return null;
                        }

                        HandleTheSameFormat(line, cfForamtDic, out activeSite, out testname, out pin, out channel, out low, out measured, out high, ref force, ref loc, ref lowUnit, ref highUnit, ref measUnit, ref number);
                        #endregion
                    }
                    else
                    {
                        #region The differ length to format header
                        string[] iA = GetlineStrArr(line);
                        if (!long.TryParse(iA[0], out long _) || !long.TryParse(iA[1], out long _))
                        {
                            return null;
                        }

                        HandleDiffFormat(line, cfForamtDic, out activeSite, out testname, out pin, out channel, ref low, ref measured, ref high, ref force, ref loc, ref lowUnit, ref highUnit, ref measUnit, ref forceUnit, out number, iA);
                        #endregion
                    }

                    #region  // unify the units
                    UnifyTheUnits(ref low, ref high, ref lowUnit, ref highUnit, ref measUnit, ref unit);
                    #endregion
                }
                #endregion

                #region Functional Type
                else // Functional 
                {
                    dataFormat = EDataFormatType.Functional;
                    string[] iA = GetlineStrArr(line);
                    if (iA.Length < 6 || !long.TryParse(iA[1], out long _) || !long.TryParse(iA[0], out long _))
                    {
                        return null;
                    }

                    activeSite = Convert.ToInt16(iA[1]);
                    measured = iA[5];
                    testname = iA[2].Trim();
                    number = Convert.ToInt64(iA[0]);
                    pat = iA[3].Trim().ToUpper().Replace(".PAT", "").Replace(".GZ", "");
                }

                #endregion

                var dataFormatRow = new DataFormatDataRow
                {
                    TestInstance = testInstance,
                    ActiveSite = activeSite,
                    DataFormatType = dataFormat,
                    Number = number,
                    TestName = testname,
                    Pin = pin,
                    Channel = channel,
                    Low = low,
                    Measured = measured,
                    High = high,
                    Force = force,
                    Loc = loc,
                    Unit = unit,
                    ForceUnit = forceUnit,
                    TestResult = testresult,
                    Pattern = pat,
                    ActuallineNumber = actuallineNum
                };
                return dataFormatRow;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private Dictionary<string, DataLogFormatHeader>? GetCfForamtDic(string line, int splitCnt, int functionHeaderMaxCnt, int parametricHeaderMaxCnt)
        {
            Dictionary<string, DataLogFormatHeader>? cfForamtDic;
            if (_currentTestNumFormat == null || _currentTestNumFormat.Count == 0 || _currentTestNumFormat["Max Length"].Length != line.Length)
            {
                cfForamtDic = _currentTestNumFormat;
                if (_lineLengthTestNumFormatParametricDict.TryGetValue(line.Length, out Dictionary<string, DataLogFormatHeader>? value))
                {
                    cfForamtDic = value;
                }

                if (splitCnt <= functionHeaderMaxCnt)
                {
                    if (_lineLengthTestNumFormatFunctionDict.TryGetValue(line.Length, out Dictionary<string, DataLogFormatHeader>? value1))
                    {
                        cfForamtDic = value1;
                    }
                }

                if (splitCnt >= parametricHeaderMaxCnt)
                {
                    if (_lineLengthTestNumFormatParametricDict.TryGetValue(line.Length, out Dictionary<string, DataLogFormatHeader>? value1))
                    {
                        cfForamtDic = value1;
                    }
                }

                _currentTestNumFormat = cfForamtDic;
            }
            else
            {
                cfForamtDic = _currentTestNumFormat;
            }

            return cfForamtDic;
        }

        private static void UnifyTheUnits(ref string low, ref string high, ref string lowUnit, ref string highUnit, ref string measUnit, ref string unit)
        {
            if (string.IsNullOrEmpty(lowUnit))
            {
                lowUnit = "N/A";
            }

            if (string.IsNullOrEmpty(highUnit))
            {
                highUnit = "N/A";
            }

            if (string.IsNullOrEmpty(measUnit))
            {
                measUnit = "N/A";
            }

            if (lowUnit != highUnit || highUnit != measUnit)
            {
                //^ + n,u,m,k,M + A/V/Ohm/Hz
                string lowUnit1 = lowUnit[..1];
                string highUnit1 = highUnit[..1];
                string measuredUnit1 = measUnit[..1];

                //change unit to align with measure unit
                if (!string.IsNullOrEmpty(low))
                {
                    low = (Convert.ToDouble(low) *
                           Math.Pow(10, UnitUtit.Instance.DicUnit[lowUnit1] - UnitUtit.Instance.DicUnit[measuredUnit1])).ToString();
                }

                if (!string.IsNullOrEmpty(high))
                {
                    high = (Convert.ToDouble(high) *
                            Math.Pow(10, UnitUtit.Instance.DicUnit[highUnit1] - UnitUtit.Instance.DicUnit[measuredUnit1])).ToString();
                }
            }

            if (measUnit != "N/A")
            {
                unit = measUnit;
            }
        }

        private static void HandleDiffFormat(string line, Dictionary<string, DataLogFormatHeader> cfForamtDic, out int activeSite, out string testname, out string pin, out string channel, ref string low, ref string measured, ref string high, ref string force, ref string loc, ref string lowUnit, ref string highUnit, ref string measUnit, ref string forceUnit, out long number, string[] iA)
        {
            int shiftNumIdx;
            int shiftIdx = 0;
            number = Convert.ToInt64(iA[0]);
            activeSite = Convert.ToInt32(iA[1]);
            testname = iA[2];
            int idxFound = Math.Max(iA[4 + shiftIdx].IndexOf('.'), iA[4 + shiftIdx].IndexOf("-1", StringComparison.Ordinal));
again:
            if (idxFound != -1)
            {
                //[4] is channel name, then it's okay to go.
            }
            else
            {
                shiftNumIdx = number.ToString().Length > cfForamtDic["Number"].Length
                    ? number.ToString().Length - cfForamtDic["Number"].Length
                    : 0;
                testname = line.Substring(cfForamtDic["Test Name"].StartIdx + shiftNumIdx, cfForamtDic["Test Name"].Length).Trim();
                testname = testname.IndexOf(' ') > -1 ? testname : iA[2];
                if (testname.Split([' '], StringSplitOptions.RemoveEmptyEntries).Length > 1)
                {
                    shiftIdx += testname.Split([' '], StringSplitOptions.RemoveEmptyEntries).Length;
                }
                else
                {
                    shiftIdx += 1;
                }
            }

            idxFound = Math.Max(iA[3 + shiftIdx].IndexOf('.'), iA[3 + shiftIdx].IndexOf("-1", StringComparison.Ordinal));
            if (idxFound != -1)
            {
                shiftIdx -= 1;
            }

            pin = idxFound == -1 ? iA[3 + shiftIdx] : "N/A";
            channel = iA.Length > 4 + shiftIdx ? iA[4 + shiftIdx] : "-1";
            idxFound = iA.Length > 4 + shiftIdx
                ? Math.Max(iA[4 + shiftIdx].IndexOf('.'), iA[4 + shiftIdx].IndexOf("-1", StringComparison.Ordinal))
                : 0;
            if (idxFound == -1 && shiftIdx <= 2)
            {
                goto again;
            }

            pin = cfForamtDic.ContainsKey("Pin") ? pin : "N/A";
            channel = cfForamtDic.ContainsKey("Channel") ? channel : "-1";
            if (!cfForamtDic.ContainsKey("Pin") || !cfForamtDic.ContainsKey("Channel"))
            {
                shiftIdx = 2;
                shiftIdx = cfForamtDic.ContainsKey("Pin") ? shiftIdx : shiftIdx - 1;
                shiftIdx = cfForamtDic.ContainsKey("Channel") ? shiftIdx : shiftIdx - 1;
            }

            int tpyeCnt = -1;
            int stAddr = 5;
            stAddr = cfForamtDic.ContainsKey("Pin") ? stAddr : stAddr - 1;
            stAddr = cfForamtDic.ContainsKey("Channel") ? stAddr : stAddr - 1;

            GetInfo(ref low, ref measured, ref high, ref force, ref loc, ref lowUnit, ref highUnit, ref measUnit, ref forceUnit, iA, shiftIdx, ref tpyeCnt, stAddr);
        }

        private static void GetInfo(ref string low, ref string measured, ref string high, ref string force, ref string loc, ref string lowUnit, ref string highUnit, ref string measUnit, ref string forceUnit, string[] iA, int shiftIdx, ref int tpyeCnt, int stAddr)
        {
            for (int idx = stAddr + shiftIdx; idx < iA.Length; idx++)
            {
                char c = iA[idx][0];
                if (iA[idx] == "N/A" || c == '0' || c == '1' || c == '2' || c == '3' || c == '4' ||
                    c == '5' ||
                    c == '6' || c == '7' || c == '8' || c == '9' || c == '-')
                // || Regex.IsMatch(iA[idx], @"^\d|-"))
                {
                    tpyeCnt++;
                    switch (tpyeCnt)
                    {
                        case 0:
                            low = iA[idx].Replace("N/A", "");
                            break;
                        case 1:
                            measured = iA[idx].Replace("N/A", "");
                            break;
                        case 2:
                            high = iA[idx].Replace("N/A", "");
                            break;
                        case 3:
                            force = iA[idx].Replace("N/A", "");
                            break;
                        case 4:
                            loc = iA[idx];
                            break;
                    }
                }
                else
                {
                    switch (tpyeCnt)
                    {
                        case 0:
                            lowUnit += iA[idx];
                            break;
                        case 1:
                            measUnit += iA[idx];
                            break;
                        case 2:
                            highUnit += iA[idx];
                            break;
                        case 3:
                            forceUnit += iA[idx];
                            force += " " + forceUnit;
                            break;
                    }
                }
            }
        }

        private static void HandleTheSameFormat(string line, Dictionary<string, DataLogFormatHeader> cfForamtDic, out int activeSite, out string testname, out string pin, out string channel, out string low, out string measured, out string high, ref string force, ref string loc, ref string lowUnit, ref string highUnit, ref string measUnit, ref long number)
        {
            string chknumber = line.Substring(cfForamtDic["Number"].StartIdx, 16).Trim();
            int shiftNumIdx = chknumber.IndexOf(' ') > cfForamtDic["Number"].Length
                  ? chknumber.IndexOf(' ') - cfForamtDic["Number"].Length
                  : 0;
            if (shiftNumIdx > 0)
            {
                number = Convert.ToInt64(line.Substring(cfForamtDic["Number"].StartIdx, cfForamtDic["Number"].Length + shiftNumIdx).Trim());
            }

            activeSite = Convert.ToInt16(line.Substring(cfForamtDic["Site"].StartIdx + shiftNumIdx, cfForamtDic["Site"].Length).Trim());

            testname = line.Substring(cfForamtDic["Test Name"].StartIdx + shiftNumIdx, cfForamtDic["Test Name"].Length).Trim();
            pin = cfForamtDic.TryGetValue("Pin", out DataLogFormatHeader? pinHeader) ? line.Substring(pinHeader.StartIdx + shiftNumIdx, pinHeader.Length).Trim() : "N/A";
            if (string.IsNullOrEmpty(pin))
            {
                pin = "N/A";
            }

            channel = "-1";
            int shiftChIdx = 0;
            if (cfForamtDic.TryGetValue("Channel", out DataLogFormatHeader? channelHeader))
            {
                channel =
                    line.Substring(channelHeader.StartIdx + shiftNumIdx, channelHeader.Length)
                        .Trim();
                string chkChannel = line.Substring(channelHeader.StartIdx + shiftNumIdx, 16);
                shiftChIdx = chkChannel.IndexOf(' ') > cfForamtDic["Channel"].Length
                    ? chkChannel.IndexOf(' ') - cfForamtDic["Channel"].Length
                    : 0;
                if (shiftChIdx > 0)
                {
                    channel =
                        line.Substring(cfForamtDic["Channel"].StartIdx + shiftNumIdx, cfForamtDic["Channel"].Length + shiftChIdx).Trim();
                }
            }
            low =
                line.Substring(cfForamtDic["Low"].StartIdx + shiftNumIdx + shiftChIdx, cfForamtDic["Low"].Length)
                    .Trim().Replace("N/A", "");
            measured =
                line.Substring(cfForamtDic["Measured"].StartIdx + shiftNumIdx + shiftChIdx, cfForamtDic["Measured"].Length).Trim().Replace("(A)", "").Replace("(F)", "");
            high = line.Substring(cfForamtDic["High"].StartIdx + shiftNumIdx + shiftChIdx, cfForamtDic["High"].Length).Trim().Replace("N/A", "");
            if (cfForamtDic.TryGetValue("Force", out DataLogFormatHeader? forceHeader))
            {
                force = line.Substring(forceHeader.StartIdx + shiftNumIdx + shiftChIdx, forceHeader.Length).Trim();
            }

            if (cfForamtDic.ContainsKey("Loc"))
            {
                loc = cfForamtDic.TryGetValue("Loc", out DataLogFormatHeader? locHeader) ? line.Substring(locHeader.StartIdx + shiftNumIdx + shiftChIdx, locHeader.Length).Trim()
                    : "0";
            }
            // -4.000 uA

            int idx = low.IndexOf(' ');
            if (idx > -1)
            {
                lowUnit = low[idx..].Trim();
                lowUnit = string.IsNullOrEmpty(lowUnit) ? "N/A" : lowUnit;
                low = low[..idx];
            }
            // 4.000 uA
            idx = high.IndexOf(' ');
            if (idx > -1)
            {
                highUnit = high[idx..].Trim();
                highUnit = string.IsNullOrEmpty(highUnit) ? "N/A" : highUnit;
                high = high[..idx];
            }
            // 0.0137 uA
            idx = measured.IndexOf(' ');
            if (idx > -1)
            {
                measUnit = measured[idx..].Trim();
                measUnit = string.IsNullOrEmpty(measUnit) ? "N/A" : measUnit;
                measured = measured[..idx];
            }
        }

        private static int GetShmooLine(string line, int activeSite, out string testname, out long number, out EDataFormatType eDataFormatType, out string pat, out string testInstance)
        {
            string shmooType = line.IndexOf(@"Y\@", StringComparison.Ordinal) > -1 ? "2D" : "1D";
            string[] isA = CommaRegex().Split(line);
            //Regex.IsMatch(line, @"Y\@") ? "2D" : "1D"; //var iA = Regex.Split(line, @"\s+");

            int resI;
            if (shmooType == "1D")
            {
                if (int.TryParse(isA[5], out resI))
                {
                    activeSite = resI;
                }
            }
            else
            {
                if (int.TryParse(isA[1], out resI))
                {
                    activeSite = resI;
                }
            }

            string patString = string.Empty;

            foreach (string s in isA)
            {
                if (!PatOrGzRegex().IsMatch(s))
                {
                    continue;
                }

                string shmooPat = BackslashPrefixRegex().Replace(s, "");
                shmooPat = PatGzSuffixRegex().Replace(shmooPat, "");
                patString += shmooPat.ToUpper() + ",";
            }

            testInstance = shmooType == "1D" ? isA[7] : isA[6];
            //Int16 : -32768 ~ +32768 不夠
            string testNum = shmooType == "1D" ? isA[9] : isA[8];
            string shmooSetup = shmooType == "1D" ? isA[8] : isA[7];
            eDataFormatType = EDataFormatType.Shmoo;
            number = Convert.ToInt64(testNum);
            testname = shmooSetup;
            pat = patString[..^1];
            return activeSite;
        }

        private static string[] GetlineStrArr(string line)
        {
            string lineDeal = line;
            if (line.IndexOf("(F)", StringComparison.Ordinal) > -1)
            {
                lineDeal = lineDeal.Replace("(F)", "");
            }

            if (line.IndexOf("(A)", StringComparison.Ordinal) > -1)
            {
                lineDeal = lineDeal.Replace("(A)", "");
            }

            string[] iA = lineDeal.Split([" "], StringSplitOptions.RemoveEmptyEntries);

            return iA;
        }
    }
}
