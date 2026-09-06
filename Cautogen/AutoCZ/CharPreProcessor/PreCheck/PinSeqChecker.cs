using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager.PatternError;
using Cautogen.AutoCZ.CharPreProcessor.Utility;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;

using CommandLine.Text;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class PinSeqChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charRows, string sheetName)
        {
            // early return if user does not input pattern info
            if (UtilityMain.UtilityData.InputParam.PatinfoFile == "")
            {
                return;
            }

            // {payload1 -> charRows that have same payload1} 
            var payload1CharRowsDict = charRows.Where(row => CharPlan.HardIpSheets.Contains(row.SheetName.ToUpper()) &&
                row.UserDef1.ToUpper() == "HAC" && row.Payload1 != "")
                .GroupBy(a => a.Payload1)
                .ToDictionary(a => a.Key.ToUpper(), a => a.ToList());

            if (payload1CharRowsDict.Count <= 0)
            {
                return;
            }

            foreach (string payload1 in payload1CharRowsDict.Keys)
            {
                List<Characterization> rowsSamePayload1 = payload1CharRowsDict[payload1];
                Characterization firstCharRow = rowsSamePayload1[0];

                if (UtilityMain.UtilityData.PatInfoDict.TryGetValue(payload1, out HardIpReference patInfo))
                {
                    var hvCharRows = payload1CharRowsDict[payload1].Where(item => item.UserDef3 == "H")
                        .Where(item => item.OnlyHasPayload1 && (item.UserDef2.ToUpper() == "MEASV" || item.UserDef2.ToUpper() == "MEASI" || item.UserDef2.ToUpper() == "MEASF"))
                            .ToList();

                    _CheckNoPatternOtherThanPayload1(hvCharRows);

                    // handle user_def6 == "multi"
                    if (hvCharRows.Count > 0 && Regex.IsMatch(hvCharRows[0].UserDef6, "multi", RegexOptions.IgnoreCase))
                    {
                        hvCharRows = hvCharRows.GroupBy(a => a.UserDef6).ToList()[0].ToList();
                    }

                    _CheckTypeOrder(hvCharRows, patInfo, firstCharRow, payload1);
                    _CheckMatchHln(rowsSamePayload1, charRows);
                    _CheckMeasCSequence(rowsSamePayload1, patInfo);
                }
                else  // can not find patternName in PatInfoDict
                {
                    PatternErrorCache.Mark(payload1, firstCharRow.SheetName, PatternErrorCache.MarkType.MissingPatternInPatInfo);
                    string outString = "Missing patten in PatInfo file in " + firstCharRow.SheetName + "  Row: " + firstCharRow.RowNum;
                    ErrorManager.AddError(ErrorType.MissingPatternInPatInfo, firstCharRow.SheetName, firstCharRow.RowNum,
                        firstCharRow.ColNum("payload1"), "Use", outString, payload1);
                    foreach (int col in firstCharRow.ColNum("payload1"))
                    {
                        ErrorReportManager.AddError(CharErrorType.E_MissingPatternInPatInfo_01, firstCharRow.SheetName, firstCharRow.RowNum, col,
                            [firstCharRow.SheetName, $"{firstCharRow.RowNum}"],
                            new ErrorInfo() { Comments = new List<string>() { payload1 } });
                    }
                }
            }
        }

        private static int PinsNum(IEnumerable<HardIpSeqInfo> infoList)
        {
            return infoList == null ? 0
                                    : infoList.Sum(info => _GetPinCount(info.MeasFPinList) +
                                        _GetPinCount(info.MeasFdiffPinList) +
                                        _GetPinCount(info.MeasIIoPinList) +
                                        _GetPinCount(info.MeasIPowerPinList) +
                                        _GetPinCount(info.MeasVIoPinList) +
                                        _GetPinCount(info.MeasVPowerPinList) +
                                        _GetPinCount(info.MeasZPinList) +
                                        _GetPinCount(info.MeasRPinList) +
                                        _GetPinCount(info.MeasVdiff2PinList) +
                                        _GetPinCount(info.MeasIdiffPinList) +
                                        _GetPinCount(info.MeasVdiffPinList));
        }

        private static void _CheckNoPatternOtherThanPayload1(IEnumerable<Characterization> charRows)
        {
            foreach (Characterization charRow in charRows)
            {
                var notNeedPatterns = new List<string>();
                if (charRow.Init1 != "")
                {
                    notNeedPatterns.Add(charRow.Init1);
                }

                if (charRow.Init2 != "")
                {
                    notNeedPatterns.Add(charRow.Init2);
                }

                if (charRow.Init3 != "")
                {
                    notNeedPatterns.Add(charRow.Init3);
                }

                if (charRow.Init4 != "")
                {
                    notNeedPatterns.Add(charRow.Init4);
                }

                if (charRow.Init5 != "")
                {
                    notNeedPatterns.Add(charRow.Init5);
                }

                if (charRow.Init6 != "")
                {
                    notNeedPatterns.Add(charRow.Init6);
                }

                if (charRow.Init7 != "")
                {
                    notNeedPatterns.Add(charRow.Init7);
                }

                if (charRow.Init8 != "")
                {
                    notNeedPatterns.Add(charRow.Init8);
                }

                if (charRow.Init9 != "")
                {
                    notNeedPatterns.Add(charRow.Init9);
                }

                if (charRow.Init10 != "")
                {
                    notNeedPatterns.Add(charRow.Init10);
                }

                if (charRow.Payload2 != "")
                {
                    notNeedPatterns.Add(charRow.Payload2);
                }

                if (charRow.Payload3 != "")
                {
                    notNeedPatterns.Add(charRow.Payload3);
                }

                if (charRow.Payload4 != "")
                {
                    notNeedPatterns.Add(charRow.Payload4);
                }

                if (charRow.Payload5 != "")
                {
                    notNeedPatterns.Add(charRow.Payload5);
                }

                if (notNeedPatterns.Count == 0)
                {
                    continue;
                }

                string outString = "There is pattern other than in payload1 in " + charRow.SheetName + "  With: " + string.Join(",", notNeedPatterns);
                ErrorManager.AddError(ErrorType.PatternOtherThanInPayload1, charRow.SheetName,
                    charRow.RowNum, charRow.ColNum(notNeedPatterns), "Use", outString);
                foreach (int col in charRow.ColNum(notNeedPatterns))
                {
                    ErrorReportManager.AddError(CharErrorType.E_PatternOtherThanInPayload1_01, charRow.SheetName, charRow.RowNum, col, [charRow.SheetName, string.Join(",", notNeedPatterns)]);
                }
            }
        }

        private static void _CheckTypeOrder(
            IList<Characterization> hvCharRows,
            HardIpReference patInfo,
            Characterization firstCharRow,
            string patternName)
        {
            if (firstCharRow == null)
            {
                return;
            }

            if (hvCharRows.Count == 0)
            {
                return;
            }

            int patMeasCount = PinsNum(patInfo.SeqInfo);
            int planMeasCount = hvCharRows.Sum(charPlan => _GetPinCount(charPlan.UserDef4));
            if (planMeasCount == patMeasCount)
            {
                // check meas sequence are the same
                if (planMeasCount <= 0)
                {
                    return;
                }

                bool flag = true;
                string message = "";
                int pinIndex = 0;
                foreach (HardIpSeqInfo info in patInfo.SeqInfo)
                {
                    _CheckMeasF(hvCharRows, info, ref pinIndex, ref message, ref flag);
                    _CheckMeasFdiff(hvCharRows, info, ref pinIndex, ref message, ref flag);
                    _CheckMeasV(hvCharRows, info, ref pinIndex, ref message, ref flag);
                    _CheckMeasI(hvCharRows, info, ref pinIndex, ref message, ref flag);
                    _CheckMeasVdiff(hvCharRows, info, ref pinIndex, ref message, ref flag);
                    _CheckMeasVdiff2(hvCharRows, info, ref pinIndex, ref message, ref flag);

                    if (flag)
                    {
                        continue;
                    }

                    UtilityMain.UtilityData.PatInfoErrorList.Add(patInfo);
                    if (message.Contains("MeasCount"))
                    {
                        PatternErrorCache.Mark(patternName, firstCharRow.SheetName, PatternErrorCache.MarkType.WrongMeasOrder);
                        string outString = message + " for HAC in " + firstCharRow.SheetName + "  Row: " + firstCharRow.RowNum;
                        ErrorManager.AddError(ErrorType.WrongMeasCount, firstCharRow.SheetName, firstCharRow.RowNum,
                            firstCharRow.ColNum("userdef4"), "Use", outString, firstCharRow.Payload1);
                        foreach (int col in firstCharRow.ColNum("userdef4"))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_WrongMeasCount_01, firstCharRow.SheetName, firstCharRow.RowNum, col,
                                [message, firstCharRow.SheetName, $"firstCharRow.RowNum"],
                                new ErrorInfo() { Comments = new List<string>() { firstCharRow.Payload1 } });
                        }
                    }
                    else
                    {
                        PatternErrorCache.Mark(patternName, firstCharRow.SheetName, PatternErrorCache.MarkType.MissingMeasPin);
                        string outString = message + " for HAC in " + firstCharRow.SheetName + "  Row: " + firstCharRow.RowNum;
                        ErrorManager.AddError(ErrorType.WrongMeasPin, firstCharRow.SheetName, firstCharRow.RowNum,
                            firstCharRow.ColNum("userdef4"), "Use", outString, firstCharRow.Payload1);
                        foreach (int col in firstCharRow.ColNum("userdef4"))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_WrongMeasPin_01, firstCharRow.SheetName, firstCharRow.RowNum, col,
                                [message, firstCharRow.SheetName, $"{firstCharRow.RowNum}"],
                                new ErrorInfo() { Comments = new List<string>() { firstCharRow.Payload1 } });
                        }
                    }
                    break;
                }
                // payload contains measseq and seq pass with check => sort pin, otherwise no action and show error in previous action
                if (patInfo.SeqInfo.Count != 0 && flag)
                {
                    //TO DO MeasType for sort criteria 
                    //pin1Pdiffpin1N
                    //pinPdiffpinN
                    //
                    foreach (HardIpSeqInfo info in patInfo.SeqInfo)
                    {
                        info.SortPinList = _SortMeasPin(info.PinList, info.SeqName);
                    }
                }
            }

            // meas count are different and pat info meas count == 0
            else if (patMeasCount == 0)
            {
                PatternErrorCache.Mark(patternName, firstCharRow.SheetName, PatternErrorCache.MarkType.MissingPinSeq);
                const string outString = "Missing pin seq or specified of pattern in PatInfo file ";
                ErrorManager.AddError(ErrorType.MissingPinSeq, firstCharRow.SheetName, firstCharRow.RowNum,
                    firstCharRow.ColNum("payload1"), "Use", outString, firstCharRow.Payload1);
                foreach (int col in firstCharRow.ColNum("payload1"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_MissingPinSeq_01, firstCharRow.SheetName, firstCharRow.RowNum, col, [],
                        new ErrorInfo() { Comments = new List<string>() { firstCharRow.Payload1 } });
                }
            }

            // measure counts are not equal
            else if (planMeasCount >= 0 && patMeasCount > 0)
            {
                UtilityMain.UtilityData.PatInfoErrorList.Add(patInfo);
                PatternErrorCache.Mark(patternName, firstCharRow.SheetName, PatternErrorCache.MarkType.WrongMeasCount);

                string outString;
                string status;

                if (planMeasCount < patMeasCount)
                {
                    outString = "Less total MeasCount in Char_Plan for HAC in " + firstCharRow.SheetName + " than in HardIP_info";
                    status = "Less";
                }
                else
                {
                    outString = "More total MeasCount in Char_Plan for HAC in " + firstCharRow.SheetName + " than in HardIP_info";
                    status = "More";
                }

                ErrorManager.AddError(ErrorType.WrongMeasCount, firstCharRow.SheetName, firstCharRow.RowNum,
                    firstCharRow.ColNum("userdef4"), "Use", outString, firstCharRow.Payload1);
                foreach (int col in firstCharRow.ColNum("userdef4"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_WrongMeasCount_02, firstCharRow.SheetName, firstCharRow.RowNum, col, [status, firstCharRow.SheetName],
                        new ErrorInfo() { Comments = new List<string>() { firstCharRow.Payload1 } });
                }
            }
        }

        private static void _CheckMeasVdiff2(IList<Characterization> newList, HardIpSeqInfo info, ref int pinIndex, ref string message, ref bool flag)
        {
            int matchCount = 0;
            foreach (string t in info.MeasVdiff2PinList)
            {
                while (newList[pinIndex].UserDef2.ToUpper() == "MEASC")
                {
                    pinIndex++;
                }

                if (newList[pinIndex].UserDef2.ToUpper() == "MEASV" &&
                    IsContainsAll(info.MeasVdiff2PinList, newList[pinIndex].UserDef4))
                {
                    matchCount += _GetPinCount(newList[pinIndex].UserDef4);
                    pinIndex++;
                    if (matchCount >= info.MeasVdiff2PinList.Count)
                    {
                        break;
                    }
                }
                else
                {
                    message = newList[pinIndex].UserDef2.ToUpper() == "MEASV"
                        ? "Wrong MeasPin of MeasVdiff2"
                        : "Wrong MeasCount of MeasVdiff2";
                    flag = false;
                    break;
                }
            }
        }

        private static void _CheckMeasVdiff(IList<Characterization> newList, HardIpSeqInfo info, ref int pinIndex, ref string message, ref bool flag)
        {
            int matchCount = 0;
            foreach (string t in info.MeasVdiffPinList)
            {
                while (newList[pinIndex].UserDef2.ToUpper() == "MEASC")
                {
                    pinIndex++;
                }
                if (newList[pinIndex].UserDef2.ToUpper() == "MEASV" &&
                    IsContainsAll(info.MeasVdiffPinList, newList[pinIndex].UserDef4))
                {
                    matchCount += _GetPinCount(newList[pinIndex].UserDef4);
                    pinIndex++;
                    if (matchCount >= info.MeasVdiffPinList.Count)
                    {
                        break;
                    }
                }
                else
                {
                    message = newList[pinIndex].UserDef2.ToUpper() == "MEASV"
                        ? "Wrong MeasPin of MeasVdiff"
                        : "Wrong MeasCount of MeasVdiff";
                    flag = false;
                    break;
                }
            }
        }

        private static void _CheckMeasFdiff(IList<Characterization> newList, HardIpSeqInfo info, ref int pinIndex, ref string message, ref bool flag)
        {
            int matchCount = 0;
            foreach (string t in info.MeasFdiffPinList)
            {
                while (newList[pinIndex].UserDef2.ToUpper() == "MEASC")
                {
                    pinIndex++;
                }
                if (newList[pinIndex].UserDef2.ToUpper() == "MEASF" &&
                    IsContainsAll(info.MeasFdiffPinList, newList[pinIndex].UserDef4))
                {
                    matchCount += _GetPinCount(newList[pinIndex].UserDef4);
                    pinIndex++;
                    if (matchCount >= info.MeasFdiffPinList.Count)
                    {
                        break;
                    }
                }
                else
                {
                    message = newList[pinIndex].UserDef2.ToUpper() == "MEASF"
                        ? "Wrong MeasPin of MeasFdiff"
                        : "Wrong MeasCount of MeasFdiff";
                    flag = false;
                    break;
                }
            }
        }

        private static void _CheckMeasI(IList<Characterization> newList, HardIpSeqInfo info, ref int pinIndex, ref string message, ref bool flag)
        {
            int matchCount = 0;
            for (int i = 0; i < info.MeasIIoPinList.Count + info.MeasIPowerPinList.Count; i++)
            {
                while (newList[pinIndex].UserDef2.ToUpper() == "MEASC")
                {
                    pinIndex++;
                }
                if (newList[pinIndex].UserDef2.ToLower() == "measi" &&
                    (IsContainsAll(info.MeasIIoPinList, newList[pinIndex].UserDef4) ||
                     IsContainsAll(info.MeasIPowerPinList, newList[pinIndex].UserDef4)))
                {
                    matchCount += _GetPinCount(newList[pinIndex].UserDef4);
                    pinIndex++;

                    if (matchCount >= info.MeasIIoPinList.Count + info.MeasIPowerPinList.Count)
                    {
                        break;
                    }
                }
                else
                {
                    message = newList[pinIndex].UserDef2.ToLower() == "measi"
                        ? "Wrong MeasPin of MeasI"
                        : "Wrong MeasCount of MeasI";
                    flag = false;
                    break;
                }
            }
        }

        private static void _CheckMeasV(IList<Characterization> newList, HardIpSeqInfo info, ref int pinIndex, ref string message, ref bool flag)
        {
            // if (info.MeasVIoPinList.Count + info.MeasVPowerPinList.Count == 0) return;

            int matchCount = 0;
            for (int i = 0; i < info.MeasVIoPinList.Count + info.MeasVPowerPinList.Count; i++)
            {
                // bypass measC
                while (newList[pinIndex].UserDef2.ToUpper() == "MEASC")
                {
                    pinIndex++;
                }
                if (newList[pinIndex].UserDef2.ToLower() == "measv" &&
                    (IsContainsAll(info.MeasVIoPinList, newList[pinIndex].UserDef4) ||
                     IsContainsAll(info.MeasVPowerPinList, newList[pinIndex].UserDef4)))
                {
                    matchCount += _GetPinCount(newList[pinIndex].UserDef4);
                    pinIndex++;
                    if (matchCount >= info.MeasVIoPinList.Count + info.MeasVPowerPinList.Count)
                    {
                        break;
                    }
                }
                else
                {
                    message = newList[pinIndex].UserDef2.ToLower() == "measv"
                        ? "Wrong MeasPin of MeasV"
                        : "Wrong MeasCount of MeasV";
                    flag = false;
                    break;
                }
            }
        }

        private static void _CheckMeasF(IList<Characterization> newList, HardIpSeqInfo info, ref int pinIndex, ref string message, ref bool flag)
        {
            // if (info.MeasFPinList.Count == 0) return;

            int matchCount = 0;

            foreach (string t in info.MeasFPinList)
            {
                // fast forward to the next one of MeasC
                while (newList[pinIndex].UserDef2.ToUpper() == "MEASC")
                {
                    pinIndex++;
                }

                if (newList[pinIndex].UserDef2.ToUpper() == "MEASF"
                    && IsContainsAll(info.MeasFPinList, newList[pinIndex].UserDef4))
                {
                    matchCount += _GetPinCount(newList[pinIndex].UserDef4);
                    pinIndex++;
                    if (matchCount >= info.MeasFPinList.Count)
                    {
                        break;
                    }
                }

                else
                {
                    message = newList[pinIndex].UserDef2.ToUpper() == "MEASF"
                        ? "Wrong MeasPin of MeasF"
                        : "Wrong MeasCount of MeasF";
                    flag = false;
                    break;
                }
            }
        }

        private static void _CheckMeasCSequence(IEnumerable<Characterization> rowsSamePayload1, HardIpReference patInfo)
        {
            // only check user_def3 = HV and user_def2 = MeasC
            var hvCharRows = rowsSamePayload1.Where(item => item.UserDef3 == "H")
                .Where(item => item.OnlyHasPayload1 && item.UserDef2.ToUpper() == "MEASC")
                .ToList();

            if (hvCharRows.Count == 0)
            {
                return;
            }

            // pre check other pattern
            _CheckNoPatternOtherThanPayload1(hvCharRows);

            var measCList = hvCharRows.Where(row => row.UserDef2.ToUpper().Contains("MEASC")).ToList();
            int patInfoWdrCount = patInfo.CapBitStr.Split('+').Count(a => a != "");

            // check plan measC count matches with pattern info
            if (measCList.Count != patInfoWdrCount)
            {
                Characterization item = hvCharRows[0];
                const string outString = "Wrong Capture count in USERDEF6";
                ErrorManager.AddError(ErrorType.WrongCapture, item.SheetName, item.RowNum, item.ColNum("userdef6"), "Use", outString, item.Payload1);
                foreach (int col in item.ColNum("userdef6"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_WrongCapture_01, item.SheetName, item.RowNum, col, []);
                }
            }

            // check the measC sequence matches with pattern info 
            else if (measCList.Count > 0)
            {
                int index = 0;
                var planCapBitNames = measCList.Select(a => a.UserDef6.ToUpper()).ToList();
                var infoCapBitNames = patInfo.CapBitStr.Split('+').Select(a => a.Split('_')[0].ToUpper()).ToList();
                bool useCapList = infoCapBitNames.Exists(planCapBitNames.Contains);

                if (!string.IsNullOrEmpty(patInfo.CapBitName) && !useCapList)
                {
                    List<string> capList = patInfo.CapBitName.ToUpper().Replace("_", "").Split('+').ToList();
                    if (capList.Distinct().ToList().Count == capList.Count)
                    {
                        infoCapBitNames = capList;
                    }
                }

                foreach (string str in infoCapBitNames)
                {
                    if (str != measCList[index].UserDef6.ToUpper())
                    {
                        Characterization item = measCList[index];
                        const string outString = "Wrong Capture sequence in USERDEF6 for MeasC";
                        ErrorManager.AddError(ErrorType.WrongCapture, item.SheetName, item.RowNum, item.ColNum("userdef6"), "Use", outString, item.Payload1, measCList[index].UserDef6, str);
                        foreach (int col in item.ColNum("userdef6"))
                        {
                            ErrorReportManager.AddError(CharErrorType.E_WrongCapture_02, item.SheetName, item.RowNum, col, []);
                        }
                    }
                    index++;
                }
            }
        }

        private static int _GetPinCount(string pinName)
        {
            if (UtilityMain.UtilityData.PinGroups.TryGetValue(pinName, out List<string> group))
            {
                return group.Count;
            }

            if (pinName.Contains("DIFF", StringComparison.OrdinalIgnoreCase) || UtilityFunction.IsCmPins(pinName))
            {
                return 2;
            }

            return pinName.Contains("::") ? 2 : 1;
        }

        private static int _GetPinCount(IEnumerable<string> pinList)
        {
            return pinList.Where(pin => pin != "").Sum(_GetPinCount);
        }

        // Return true if listB is in listA
        private static bool IsContainsAll(IEnumerable<string> listA, string planPin)
        {
            var listC = new List<string>();
            var listD = new List<string>();

            foreach (string pinName in listA.Select(pin => pin.Replace("_", "")))
            {
                if (pinName.Contains("::")) //PinA_P::PinA_N ---- Diff Pins
                {
                    string pinName1 = pinName.Split(':')[0];
                    string pinName2 = pinName.Split(':')[2];
                    listC.Add(pinName1);
                    listC.Add(pinName2);
                    continue;
                }

                if (UtilityMain.UtilityData.PinGroupList.TryGetValue(pinName, out List<string> value))
                {
                    listC.AddRange(value);
                }
                else
                {
                    listC.Add(pinName);
                }
            }

            // for Common mode output swing, the pin name in usdef4 is pin1CMpin2.
            if (UtilityFunction.IsCmPins(planPin))
            {
                string[] array = Regex.Split(planPin, "CM", RegexOptions.IgnoreCase);
                listD.AddRange(array);
            }

            // PCIETX0VDIFF------Diff Pins
            else if (Regex.IsMatch(planPin, "Diff$", RegexOptions.IgnoreCase))
            {
                int length = planPin.Length;
                string pinName1 = planPin.Substring(0, ((length - 5) / 2) - 1);
                string pinName2 = planPin.Substring((length - 5) / 2, ((length - 5) / 2) - 1);
                if (pinName1 == pinName2)
                {
                    listD.Add(pinName1 + "P");
                    listD.Add(pinName1 + "N");

                }
                else
                {
                    listD.Add(planPin.Substring(0, length - 5) + "P");
                    listD.Add(planPin.Substring(0, length - 5) + "N");
                }
            }
            // PCIETX0PDIFFPCIETX0N---- Diff Pins
            else if (Regex.IsMatch(planPin, "[a-zA-Z]+DIFF[a-zA-Z]+", RegexOptions.IgnoreCase))
            {
                string[] array = Regex.Split(planPin, "DIFF", RegexOptions.IgnoreCase);
                listD.AddRange(array);
            }
            else if (UtilityMain.UtilityData.PinGroupList.TryGetValue(planPin, out List<string> value))
            {
                listD.AddRange(value);
            }
            else
            {
                listD.Add(planPin);
            }

            return listD.All(d => listC.Any(c => c == d));
        }

        private static void _CheckMatchHln(IList<Characterization> charList, ICollection<Characterization> allCharRowsInSheet)
        {
            bool isMatchHln = true;
            var tpNames = (from row in charList select row.TpName).Distinct().ToList();

            int charCount = charList.Count;  // cache the charList.count since it will change length in the loop
            for (int i = 0; i < charCount; i++)
            {
                if (charList[i].UserDef3 == "H")
                {
                    _UpdateHvTestName(charList[i]);
                }

                _UpdateCounterRow(allCharRowsInSheet, charList, tpNames, i, ref isMatchHln, "H");
                _UpdateCounterRow(allCharRowsInSheet, charList, tpNames, i, ref isMatchHln, "L");
                _UpdateCounterRow(allCharRowsInSheet, charList, tpNames, i, ref isMatchHln, "N");
                _UpdateCounterRow(allCharRowsInSheet, charList, tpNames, i, ref isMatchHln, "N1");
            }

            if (isMatchHln)
            {
                return;
            }

            Characterization firstCharRow = charList[0];
            string outString = "Missing HLN condition in " + firstCharRow.SheetName;
            ErrorManager.AddWarning(ErrorType.MissingHlnCondition, firstCharRow.SheetName, firstCharRow.RowNum,
                firstCharRow.ColNum("userdef3"), "Use", outString, firstCharRow.Payload1);
            foreach (int col in firstCharRow.ColNum("userdef3"))
            {
                ErrorReportManager.AddError(CharErrorType.W_MissingHlnCondition_01, firstCharRow.SheetName, firstCharRow.RowNum, col, [firstCharRow.SheetName]);
            }
        }

        private static void _UpdateCounterRow(ICollection<Characterization> allCharRowsInSheet, IList<Characterization> charList, ICollection<string> tpNames, int i, ref bool isMatchHln, string hln)
        {
            // early leave if no any hln row
            if (charList.All(row => row.UserDef3 != hln))
            {
                return;
            }

            // early leave if has hln counter row
            if (tpNames.Contains(_ReplaceTpName(charList[i].TpName, hln)))
            {
                return;
            }

            // flag has no hln counter row
            isMatchHln = false;

            // create counter row and push into lists
            Characterization newChar = _GetHlnCounterRow(charList[i], hln, hln == "N1" ? "NV" : hln + "V");
            charList.Add(newChar);
            allCharRowsInSheet.Add(newChar);
            tpNames.Add(newChar.TpName);
            newChar.TpName = newChar.TpName.Replace(newChar.UserDef1 + "_", "TSMC_"); // can not merge into _GetHlnCounterRow

            // update hv test name
            if (hln == "H")
            {
                _UpdateHvTestName(newChar);
            }
        }

        private static Characterization _GetHlnCounterRow(Characterization charRow, string targetVoltage, string otherVoltage)
        {
            return new Characterization(charRow)
            {
                UserDef3 = targetVoltage,
                OtherSupplies = otherVoltage,
                TpName = _ReplaceTpName(charRow.TpName, targetVoltage),
                IpUse1 = _ReplaceTpName(charRow.IpUse1, targetVoltage)
            };
        }

        private static void _UpdateHvTestName(Characterization hvCharRow)
        {
            // added on 10/6, to sort Test Names
            string key = hvCharRow.SheetName + "&" + hvCharRow.Payload1;

            if (UtilityMain.UtilityData.HTestNames.ContainsKey(key))
            {
                UtilityMain.UtilityData.HTestNames[key] += "," + hvCharRow.TpName;
            }
            else
            {
                UtilityMain.UtilityData.HTestNames.Add(key, hvCharRow.TpName);
            }
        }

        private static string _ReplaceTpName(string oriTpName, string targetVoltage)
        {
            return oriTpName
                .Replace("_H_", "_" + targetVoltage + "_")
                .Replace("_N_", "_" + targetVoltage + "_")
                .Replace("_L_", "_" + targetVoltage + "_")
                .Replace("_N1_", "_" + targetVoltage + "_");
        }

        private static string _SortMeasPin(string pinList, string measType)
        {
            var pins = Regex.Split(pinList, ",").Where(s => !string.IsNullOrEmpty(s)).ToList();
            var temp2Sort = new List<string>();
            var vod2Sort = new List<string>();

            foreach (string pin in pins)
            {
                KeyValuePair<string, string> originPin =
                    PatInfoReader.PinUnderlineDict.FirstOrDefault(
                        p => p.Value.Equals(pin, StringComparison.OrdinalIgnoreCase));
                //before sort pin, restore to contains "_" pin name

                if (originPin.Key != null) //single-end
                {
                    temp2Sort.Add(originPin.Key);
                }
                else //not single-end
                {
                    if (pin.Contains("::"))
                    {
                        if (measType == "FDIFF")
                        {
                            temp2Sort.Add(pin);
                        }
                        else //
                        {
                            string pinName1 = pin.Split(':')[0];
                            string pinName2 = pin.Split(':')[2];
                            string pindiff = pin.Replace("::", "DIFF");

                            temp2Sort.Add(pinName1);
                            temp2Sort.Add(pinName2);
                            vod2Sort.Add(pindiff);
                        }
                    }
                    else
                    {
                        temp2Sort.Add(originPin.Key);
                    }
                }
            }

            vod2Sort.Sort(delegate (string x, string y)
            {
                string a = x;
                a = a.Split(':')[0];
                string b = y;
                b = b.Split(':')[0];
                return string.CompareOrdinal(a, b);
            });
            var vcm = vod2Sort.Select(vod => vod.Replace("DIFF", "VCM")).ToList();

            //sort...................
            temp2Sort.Sort(delegate (string x, string y)
            {
                string a = x;
                string b = y;
                if (a.Contains("::"))
                {
                    a = a.Split(':')[0];
                }

                if (b.Contains("::"))
                {
                    b = b.Split(':')[0];
                }

                return string.CompareOrdinal(a, b);
            });

            temp2Sort.AddRange(vod2Sort);
            temp2Sort.AddRange(vcm);
            return string.Join(",", temp2Sort).Replace("_", "");
        }
    }
}
