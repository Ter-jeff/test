using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;
using LogLib.Utility;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class PinSeqChecker : HardIpPrecheckBase
    {
        private HardIpInfo _patInfo;
        internal List<string> _seqlstInMisc;


        public PinSeqChecker(HardIpSheet hardIpSheet, HardIpPattern pattern) : base(hardIpSheet, pattern)
        {
        }

        internal void CheckSeq(HardIpPattern patternItem, HardIpInfo patInfo)
        {
            var patternPins = patternItem.MeasPins.Where(p => !string.IsNullOrEmpty(p.MeasType) &&
                !(p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.MeasCalcLimit, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.MeasC, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.MeasLimit, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.MeasVocm, StringComparison.OrdinalIgnoreCase)
                )).ToList();
            var tmpPinList = new List<MeasPin>();
            foreach (MeasPin pin in patternPins)
            {
                if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(pin.PinName))
                {
                    foreach (string groupPin in TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(pin.PinName))
                    {
                        var newpin = new MeasPin();
                        newpin.Copy(pin);
                        newpin.PinName = groupPin;
                        newpin.SequenceIndex = pin.SequenceIndex;
                        tmpPinList.Add(newpin);
                    }
                }
                else
                {
                    tmpPinList.Add(pin);
                }
            }
            patternPins = tmpPinList;
            string message = "";
            int measIndex = HardIpSheet.PlanHeaderIdx["measIndex"];
            for (int seqIndex = 0; seqIndex < patInfo.SeqInfo.Count; seqIndex++)
            {
                //if one measure sequence is ignored, should not define limits for this measure sequence in testPlan
                if (_seqlstInMisc != null && (seqIndex >= _seqlstInMisc.Count || _seqlstInMisc[seqIndex].Equals("N", StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                HardIpSeqInfo info = patInfo.SeqInfo[seqIndex];

                if (info.PinList.Count > 0)
                {
                    foreach (string pinName in info.PinList)
                    {
                        MatchSeqPin(pinName, info, patternItem, patternPins, measIndex);
                    }
                }
            }
            foreach (MeasPin planPin in patternPins)
            {
                message = $"this pin :{planPin.PinName} can not find in pattern info";
                ErrorReportManager.AddError(
                    HardIpErrorType.W_MissingTestplanPin_01,
                    patternItem.SheetName,
                    planPin.RowNum,
                    measIndex,
                    [planPin.PinName, patInfo.Version]
                );
            }
        }

        private void MatchSeqPin(string pinName, HardIpSeqInfo info, HardIpPattern patternItem, List<MeasPin> patternPins, int measIndex)
        {
            MeasPin targetPin = patternPins.FirstOrDefault(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase) &&
                                                                p.MeasType.Equals("Meas" + info.SeqName, StringComparison.OrdinalIgnoreCase));
            if (targetPin == null && pinName.Contains("::"))
            {
                string[] pinlist = Regex.Split(pinName, "::", RegexOptions.IgnoreCase);
                string revDiffpin = string.Join("::", new List<string> { pinlist[1], pinlist[0] });
                targetPin = patternPins.FirstOrDefault(p => p.PinName.Equals(revDiffpin, StringComparison.OrdinalIgnoreCase) &&
                p.MeasType.Equals("Meas" + info.SeqName, StringComparison.OrdinalIgnoreCase));
            }
            if (targetPin == null && pinName.Contains("="))
            {
                targetPin = patternPins.FirstOrDefault(p => p.PinName.Equals(pinName.Split('=')[1], StringComparison.OrdinalIgnoreCase) &&
                p.MeasType.Equals("Meas" + info.SeqName, StringComparison.OrdinalIgnoreCase));
            }


            if (targetPin == null)
            {
                string message = $"pattern info pin:{pinName}, can't match to test plan";
                ErrorReportManager.AddError(
                    HardIpErrorType.E_WrongMeasPinInPatInfo_02,
                    patternItem.SheetName,
                    patternItem.RowNum,
                    measIndex,
                    [pinName]
                );
            }
            else
            {
                patternPins.Remove(targetPin);
                if (targetPin.PinName.Contains("::") && targetPin.MeasType == MeasType.MeasVdiff)
                {
                    string[] pinlist = Regex.Split(pinName, "::", RegexOptions.IgnoreCase);
                    foreach (string pin in pinlist)
                    {
                        targetPin = patternPins.FirstOrDefault(p => p.PinName.Equals(pin, StringComparison.OrdinalIgnoreCase) &&
                p.MeasType.Equals(MeasType.MeasV, StringComparison.OrdinalIgnoreCase));
                        if (targetPin != null)
                        {
                            patternPins.Remove(targetPin);
                        }
                    }
                }

            }
        }

        #region Sub utility Functions
        /// <summary>
        /// Get measure counts of each pattern from pattern info file
        /// </summary>
        /// <param name="infoList"></param>
        /// <returns></returns>
        internal int PinsNum_PatInfo(List<HardIpSeqInfo> infoList, List<string> seqlstInMisc)
        {
            int count = 0;
            for (int seqIndex = 0; seqIndex < infoList.Count; seqIndex++)
            {
                HardIpSeqInfo info = infoList[seqIndex];
                if (seqlstInMisc != null && (seqIndex >= seqlstInMisc.Count || seqlstInMisc[seqIndex].Equals("N", StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                //Idiff(3X):  pin_P::pin_N ----> (I)pin_P,pin_N,(Idiff)pin_P::pin_N 
                //Vdiff(4X):  pin_P::pin_N ----> (V)pin_P,pin_N,(Vdiff)pin_P::pin_N,(Vocm)pin_P::pin_N
                count += GetPinNum(info.PinList);
            }
            return count;
        }

        /// <summary>
        /// Get measure counts of each pattern from pattern info file
        /// </summary>
        /// <returns></returns>
        public int PinsNum_Plan(HardIpPattern pattern)
        {
            int count = 0;
            int totalCnt = 0;
            int vdiffCnt = 0;
            bool isMeasVdiff = SearchInfo.IsMeasVdiff(pattern);
            if (isMeasVdiff)
            {
                foreach (MeasPin pin in pattern.MeasPins)
                {
                    //excluding count of measvdiff & measvocm
                    if (pin.MeasType.ToLower() == "measv" || pin.MeasType.ToLower() == "measi" ||
                       pin.MeasType.ToLower() == "measf" || pin.MeasType.ToLower() == "measr1" ||
                       pin.MeasType.ToLower() == "measr2" || pin.MeasType.ToLower() == "measvdiff2")
                    {
                        totalCnt += pin.PinCount;
                    }

                    if (pin.MeasType.ToLower() == "measvdiff")
                    {
                        List<string> nameList = Regex.Split(pin.PinName.ToLower(), "::").ToList();
                        foreach (string pinName in nameList)
                        {
                            if (!pattern.MeasPins.Where(y => y.MeasType == MeasType.MeasV).Select(x => x.PinName.ToLower()).Contains(pinName))
                            {
                                vdiffCnt++;
                            }
                        }
                    }
                }
                count = totalCnt + vdiffCnt;
            }
            else
            {
                foreach (MeasPin pin in pattern.MeasPins)
                {
                    if (pin.MeasType.ToLower() == "measv" || pin.MeasType.ToLower() == "measi" ||
                        pin.MeasType.ToLower() == "measf" || pin.MeasType.ToLower() == "measr1" ||
                        pin.MeasType.ToLower() == "measr2" || pin.MeasType.ToLower() == "measvdiff2")
                    {
                        count += pin.PinCount;
                    }
                }
            }

            return count;
        }

        private int GetPinNum(List<string> pinList)
        {
            int count = 0;
            foreach (string pin in pinList)
            {
                if (pin == "")
                {
                    continue;
                }

                if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(pin))
                {
                    List<string> pinGroupList = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(pin);
                    count += pinGroupList.Count;
                }
                else
                {
                    if (pin.Contains("::"))
                    {
                        count += 2;
                    }
                    else
                    {
                        count++;
                    }
                }
            }
            return count;
        }

        private List<string> GetPinList(HardIpPattern pattern)
        {
            var pinlist = new List<string>();
            var measPins = pattern.MeasPins.Where(p => p.SequenceIndex > 0).SelectMany(p => Regex.Split(p.PinName, "::|,", RegexOptions.IgnoreCase)).ToList();
            foreach (string pin in measPins)
            {
                foreach (string pinGroupPin in TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(pin))
                {
                    pinlist.AddRange(DataConvertor.ConvertToNetName(pinGroupPin, TestPlanStatic.PowerMergeSheet.PowerMerge).Split(','));

                }
            }

            return pinlist;
        }

        private List<string> GetPinList(HardIpInfo info)
        {
            IEnumerable<List<string>> allTypeList = info.SeqInfo.Select(p => p.PinList);

            return allTypeList.SelectMany(p => p).Where(p => !string.IsNullOrEmpty(p))
                .SelectMany(p => Regex.Split(p, "::|,", RegexOptions.IgnoreCase)).ToList();
        }

        public int CapBitsInTp(HardIpPattern pattern)
        {
            int capBits = 0;
            foreach (MeasPin measPin in pattern.MeasPins.Where(p => p.MeasType.Equals(MeasType.MeasC)))
            {
                if (!string.IsNullOrEmpty(measPin.CapBit))
                {
                    capBits += int.Parse(measPin.CapBit);
                }
                else
                {
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongMeasC_02,
                        pattern.SheetName,
                        measPin.RowNum,
                        0,
                        []
                    );
                }
            }
            return capBits;
        }

        #endregion

        public override void Check()
        {
            try
            {
                _patInfo = HardIpService.GetHardIpInfo(Pattern);
                _seqlstInMisc = SearchInfo.GetSeqlstFromMiscInfo(Pattern.MiscInfo);
                bool ignorePatInfo = false;
                if (Pattern.Pattern.RealPatternName.Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase) ||
                    HardIpConstData.RegInsInPatt.IsMatch(Pattern.Pattern.RealPatternName) ||
                    HardIpConstData.RegOpcodeInPatt.IsMatch(Pattern.Pattern.RealPatternName))
                {
                    return;
                }
                if (Regex.IsMatch(Pattern.MiscInfo, HardIpConstData.IgnorePatInfo, RegexOptions.IgnoreCase))
                {
                    ignorePatInfo = true;
                    Response.Report("Will ignore measurements and pattern comment in patInfo, to generate test instance only according to testPlan for '" +
                                    Pattern.Pattern.GetLastPayload() + "'!", EnumMessageLevel.Warning, 20);
                }
                string errorMessage = "";

                if (CapBitsInTp(Pattern) != _patInfo.CapBit && !Pattern.Pattern.IsMultiTimeDomain())
                {
                    errorMessage =
                        $"MeasC capture bit mismatch between test plan {CapBitsInTp(Pattern)} and patInfo {_patInfo.CapBit}";
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_WrongMeasC_03,
                        Pattern.SheetName,
                        Pattern.RowNum,
                        0,
                        [$"{CapBitsInTp(Pattern)}", $"{_patInfo.CapBit}"]
                    );
                }

                if (_patInfo.Payload == Pattern.Pattern.GetLastPayload() && _patInfo.PatInfoExist)
                {
                    #region When "ignore_patt_comment", check total number of measureSequence,send bit number and cap bit number between testplan and patInfo
                    if (ignorePatInfo)
                    {
                        return;
                    }
                    #endregion

                    #region Check Measure Sequence

                    int patMeasCount = PinsNum_PatInfo(_patInfo.SeqInfo, _seqlstInMisc);
                    int planMeasCount = PinsNum_Plan(Pattern);
                    //SearchInfo.GetPlanPinsNum(patternItem);
                    if (planMeasCount == patMeasCount && patMeasCount > 0)
                    {
                        //Check test plan Meas Sequence with patInfo. 
                        CheckSeq(Pattern, _patInfo);
                    }
                    else if (planMeasCount > 0 && patMeasCount > 0 && planMeasCount != patMeasCount)
                    {
                        if (planMeasCount != patMeasCount)
                        {
                            errorMessage = "Total MeasCount different";
                        }

                        try
                        {
                            List<string> planPins = DecomposePinGroup(GetPinList(Pattern));
                            List<string> infoPins = DecomposePinGroup(GetPinList(_patInfo));
                            if (planPins.Count != infoPins.Count ||
                                (planPins.Except(infoPins).Any() && infoPins.Except(planPins).Any()))
                            {
                                string infoPinsStr = string.Join(",", planPins.Except(infoPins));
                                string planPinsStr = string.Join(",", infoPins.Except(planPins));
                                ErrorReportManager.AddError(
                                    HardIpErrorType.W_WrongTotalMeasCount_01,
                                    Pattern.SheetName,
                                    Pattern.RowNum,
                                    0,
                                    [infoPinsStr, planPinsStr]
                                );
                            }
                        }
                        catch (Exception ex)
                        {
                            ErrorMessageBox.Show(string.Format(ex.ToString()));
                        }
                    }
                    else if (planMeasCount > 0 && patMeasCount == 0)
                    {
                        errorMessage = "No measure sequence in PatInfo file, but specified measure pins in testPlan";
                        ErrorReportManager.AddError(
                            HardIpErrorType.W_MissingPinSeq_01,
                            Pattern.SheetName,
                            Pattern.RowNum,
                            0,
                            []
                        );
                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }

        }

        private List<string> DecomposePinGroup(List<string> pinsList)
        {
            var result = new List<string>();

            foreach (string pin in pinsList)
            {
                result.AddRange(TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(pin));
            }

            return result;
        }
    }
}
