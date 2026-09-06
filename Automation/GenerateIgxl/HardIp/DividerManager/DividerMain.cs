using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Utility.Basic;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using LogLib.Static;
using LogLib.Utility;

using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.DividerManager
{
    public class DividerMain
    {
        private static HardIpInfo _patInfo;
        private static bool _isMeasVdiff2;
        private static bool _isMeasVdiff;
        private static bool _isHardIpmtdTest;

        public static List<HardIpPattern> DivideInstancePattern(HardIpInputData hardIpInputData, List<HardIpPattern> patterns)
        {
            var resuList = new List<HardIpPattern>();
            foreach (HardIpPattern pattern in patterns)
            {
                try
                {
                    //If “Instance :XXX” in the pattern column but no “calc” specified in the “misc info” and no measureSeq in testPlan, will not generate instance
                    if ((HardIpConstData.RegInsInPatt.IsMatch(pattern.Pattern.RealPatternName) &&
                         pattern.MeasPins.Count == 0 && pattern.MiscInfo == "") ||
                         HardIpConstData.RegOpcodeInPatt.IsMatch(pattern.Pattern.RealPatternName))
                    {
                        continue;
                    }

                    pattern.FunctionName = pattern.FunctionName == "" ? SearchInfo.GetVbtNameByPattern(hardIpInputData, pattern) : pattern.FunctionName;
                    bool isHardIpUniversal = pattern.FunctionName == "";
                    pattern.SkipDotNet = JudgeSkipForDotNet(pattern); //Skip item for dot net test program.

                    List<HardIpPattern> tempList1 = DivideMeasPins(hardIpInputData, pattern, isHardIpUniversal);
                    if (tempList1.Count == 0)
                    {
                        string errorMessage = $"Generating Instance Error: Sheet: {pattern.SheetName}, Row Num: {pattern.RowNum}, {pattern.Pattern.TestPlanPatternName} can not find in PatInfo file!";
                        Response.Report(errorMessage, EnumMessageLevel.Error, 30);
                    }

                    resuList.AddRange(tempList1);
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                    throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum);
                }
            }

            return resuList;
        }

        public static List<HardIpPattern> DivideMeasPins(HardIpInputData hardIpInputData, HardIpPattern pattern, bool isHardIpUniversal, bool isFlowUse = false)
        {
            var resultList = new List<HardIpPattern>();
            _patInfo = HardIpService.GetHardIpInfo(pattern);

            SearchInfo.InitialMeasC(ref pattern, _patInfo);
            _isHardIpmtdTest = pattern.Pattern.IsMultiTimeDomain();
            _isMeasVdiff2 = SearchInfo.IsMeasVdiff2(pattern);
            _isMeasVdiff = SearchInfo.IsMeasVdiff(pattern);
            bool isRepeatLimit = SearchInfo.IsRepeatLimit(pattern.MiscInfo);
            var fPins = new List<MeasPin>();
            var wiPins = new List<MeasPin>();

            var totalPins = new List<MeasPin>();
            var patternVir = new HardIpPattern();
            var patternFreq = new HardIpPattern();
            var otherPins = new List<MeasPin>();

            foreach (MeasPin pin in pattern.MeasPins)
            {
                if (pin.MeasType == MeasType.MeasCalc || pin.MeasType == MeasType.MeasLimit || pin.MeasType == MeasType.MeasC || pin.MeasType == MeasType.MeasN || pin.MeasType == "")
                {
                    otherPins.Add(pin);
                }
            }

            #region Test plan does not have any Meas Pins, refer to pattern info. If pattern info contains src or capture information, choose VFI as VBT. otherwise, Function T.

            bool isUseTestPlan = false;
            if (_patInfo.SeqInfo.Count == 0)
            {
                isUseTestPlan = true;
                if (pattern.MeasPins.Count == 0)
                {
                    if (isHardIpUniversal)
                    {
                        if (_isHardIpmtdTest)
                        {
                            pattern.FunctionName = VbtFunctionLibShared.HardIpmtdTest;
                        }
                        else
                        {
                            pattern.FunctionName = VbtFunctionLibShared.VifName; // "Meas_FreqVoltCurr_Universal_func";
                        }
                    }

                    resultList.Add(pattern);
                    return resultList;
                }

                RemoveMeasPinsByMeasType(pattern);
                totalPins.AddRange(pattern.MeasPins);
            }
            else
            {
                CollectTotalPinsFromSeqInfo(totalPins);
            }

            #endregion

            #region Align measpins with VFI and VIR

            fPins = totalPins.Where(p => p.MeasType.Equals(MeasType.MeasF) || p.MeasType.Equals(MeasType.MeasFdiff) || p.MeasType.Equals(MeasType.MeasN) || p.MeasType.Equals(MeasType.MeasD)).ToList();

            wiPins = totalPins.Where(p => p.MeasType.Equals(MeasType.WiSrc) ||
                              p.MeasType.Equals(MeasType.WiMeas)).ToList();
            if (isHardIpUniversal)
            {
                fPins = totalPins;
            }

            #endregion

            totalPins.RemoveAll(otherPins.Contains);
            if (!isHardIpUniversal)
            {
                resultList.Add(BuildNonUniversalPattern(pattern, totalPins, otherPins, isUseTestPlan, isFlowUse, isRepeatLimit));
                return resultList;
            }

            if (!_isMeasVdiff2)
            {
                //virflag, freqflag determine VBT function to use
                //if exist MeasR1/R2, need to use VIR => virflag should be true
                //if exist MeasF, need to use VFI => freqflag should be true
                //non of above, use VFI

                #region BB RF part

                if (wiPins.Count > 0)
                {
                    resultList.Add(BuildRfbbPattern(pattern, wiPins, isUseTestPlan, isFlowUse, isRepeatLimit));
                    return resultList;
                }

                #endregion

                AddFreqOrOtherPins(resultList, pattern, fPins, otherPins, patternFreq, patternVir, isUseTestPlan, isFlowUse, isRepeatLimit);
            }
            #region Vdiff2 as Meas_Lpdp_Vdiff2_fuc
            else
            {
                pattern.FunctionName = VbtFunctionLibShared.VifName;
                resultList.Add(pattern);
            }
            #endregion

            return resultList;
        }

        private static HardIpPattern BuildRfbbPattern(HardIpPattern pattern, List<MeasPin> wiPins, bool isUseTestPlan, bool isFlowUse, bool isRepeatLimit)
        {
            var rfbbPattern = new HardIpPattern();
            var sortedPins = new List<MeasPin>(wiPins);
            sortedPins = ProcessSpecialMeasItems(sortedPins, _isMeasVdiff2, _isMeasVdiff, isUseTestPlan, isFlowUse);
            if (pattern.TestPlanSequences.Count != _patInfo.SeqInfo.Count)
            {
                SearchInfo.GetPlanCurrentRange(pattern.MeasPins, sortedPins, isRepeatLimit);
            }
            else
            {
                sortedPins = new List<MeasPin>(pattern.MeasPins);
            }

            rfbbPattern.Copy(pattern);
            rfbbPattern.MeasPins = sortedPins;
            rfbbPattern.FunctionName = _isHardIpmtdTest ? VbtFunctionLibShared.HardIpmtdTest : pattern.FunctionName;

            return rfbbPattern;
        }

        private static void AddFreqOrOtherPins(List<HardIpPattern> resultList, HardIpPattern pattern, List<MeasPin> fPins, List<MeasPin> otherPins, HardIpPattern patternFreq, HardIpPattern patternVir, bool isUseTestPlan, bool isFlowUse, bool isRepeatLimit)
        {
            #region freq(default use)

            if (fPins.Count > 0)
            {
                if (pattern.TestPlanSequences.Count != _patInfo.SeqInfo.Count)
                {
                    SearchInfo.GetPlanCurrentRange(pattern.MeasPins, fPins, isRepeatLimit);
                }
                else
                {
                    fPins = new List<MeasPin>(pattern.MeasPins);
                }
                fPins = ProcessPinsWithPowerMerge(fPins);
                fPins = ProcessSpecialMeasItems(fPins, _isMeasVdiff2, _isMeasVdiff, isUseTestPlan, isFlowUse);
                var sortedPins = new List<MeasPin>(fPins);

                patternFreq.Copy(pattern);

                patternFreq.MeasPins = sortedPins;
                patternFreq.FunctionName = _isHardIpmtdTest ? VbtFunctionLibShared.HardIpmtdTest : VbtFunctionLibShared.VifName; //"Meas_FreqVoltCurr_Universal_func";
                RemoveMeasPinsByMeasType(patternVir);

                if (_isMeasVdiff)
                {
                    patternFreq.SpecialMeasType =
                        pattern.MeasPins.Exists(
                            s => s.MeasType.Equals(MeasType.MeasVocm, StringComparison.OrdinalIgnoreCase))
                            ? MeasType.MeasVocm
                            : MeasType.MeasVdiff;
                }

                //DividerCommonLogic.RemoveIgnoredSequence(patternFreq);
                patternFreq.MeasPins.AddRange(otherPins);
                resultList.Add(patternFreq);
            }

            #endregion

            #region other pins

            else if (otherPins.Count > 0)
            {
                patternFreq.Copy(pattern);
                patternFreq.MeasPins = otherPins;
                patternFreq.FunctionName = _isHardIpmtdTest ? VbtFunctionLibShared.HardIpmtdTest : VbtFunctionLibShared.VifName; //"Meas_FreqVoltCurr_Universal_func";
                resultList.Add(patternFreq);
            }

            #endregion
        }

        private static void CollectTotalPinsFromSeqInfo(List<MeasPin> totalPins)
        {
            if (_patInfo.NewInfo != null)
            {
                foreach (HardIpSeqInfoNew info in _patInfo.NewInfo.SeqInfo)
                {
                    totalPins.AddRange(info.MeasPins);
                }
            }
            else
            {
                foreach (HardIpSeqInfo info in _patInfo.SeqInfo)
                {
                    foreach (MeasPin measPin in info.MeasPins)
                    {
                        if (measPin.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase))
                        {
                            List<string> pnPins = Regex.Split(measPin.PinName, "::").ToList();
                            pnPins.Sort();
                            foreach (string pin in pnPins)
                            {
                                var copyseqPin = new MeasPin();
                                copyseqPin.Copy(measPin);
                                copyseqPin.PinName = pin;
                                copyseqPin.MeasType = MeasType.MeasV;
                                copyseqPin.SequenceIndex = measPin.SequenceIndex;
                                totalPins.Add(copyseqPin);
                            }
                        }
                        totalPins.Add(measPin);
                    }
                }
            }
        }

        private static HardIpPattern BuildNonUniversalPattern(HardIpPattern pattern, List<MeasPin> totalPins, List<MeasPin> otherPins, bool isUseTestPlan, bool isFlowUse, bool isRepeatLimit)
        {
            var customVbtPattern = new HardIpPattern();
            if (!pattern.FunctionName.Equals(VbtFunctionLibShared.RfFunc) &&
                !pattern.FunctionName.Equals(VbtFunctionLibShared.RfTrim) &&
                !pattern.FunctionName.Equals(VbtFunctionLibShared.RfTrim2D) &&
                !pattern.FunctionName.Equals(VbtFunctionLibShared.LcdMeas))
            {
                var sortedPins = new List<MeasPin>(totalPins);
                sortedPins = ProcessSpecialMeasItems(sortedPins, _isMeasVdiff2, _isMeasVdiff, isUseTestPlan, isFlowUse);
                if (pattern.TestPlanSequences.Count != _patInfo.SeqInfo.Count)
                {
                    SearchInfo.GetPlanCurrentRange(pattern.MeasPins, sortedPins, isRepeatLimit);
                }
                else
                {
                    sortedPins = new List<MeasPin>(pattern.MeasPins);
                    sortedPins.RemoveAll(p => otherPins.Any(q => q.RowNum == p.RowNum));
                }
                customVbtPattern.Copy(pattern);
                customVbtPattern.MeasPins = sortedPins;
                customVbtPattern.FunctionName = pattern.FunctionName;
                customVbtPattern.MeasPins.AddRange(otherPins);
                return customVbtPattern;
            }

            return pattern;
        }

        private static List<MeasPin> ProcessPinsWithPowerMerge(List<MeasPin> allPins)
        {
            var result = new List<MeasPin>();
            foreach (MeasPin pin in allPins)
            {
                if (TestProgram.IgxlWorkBk.PinMapPair.Value.TryGetGroup(pin.PinName, out PinGroup group) &&
                    !pin.MiscInfo.ContainsIgnoreCase(HardIpConstData.IgnoreSplitPOWERPowerMerge))
                {
                    var tmpPins = new List<string>();
                    var pins = group.PinList.Select(x => x.PinName).ToList();
                    foreach (string groupPin in pins)
                    {
                        tmpPins.AddRange(DataConvertor.ConvertToNetName(groupPin, TestPlanStatic.PowerMergeSheet.PowerMerge).Split(','));
                    }
                    if (tmpPins.SequenceEqual(pins))
                    {
                        result.Add(pin);
                    }
                    else
                    {
                        foreach (string groupPin in tmpPins)
                        {
                            var copyTargetPin = new MeasPin
                            {
                                PinName = groupPin,
                                SequenceIndex = pin.SequenceIndex
                            };
                            copyTargetPin.Copy(pin);
                            result.Add(copyTargetPin);
                        }
                    }
                }
                else
                {
                    result.Add(pin);
                }
            }
            return result;
        }

        private static void RemoveMeasPinsByMeasType(HardIpPattern pattern)
        {
            IEnumerable<IGrouping<int, MeasPin>> pinGroups = from pin in pattern.MeasPins
                                                             group pin by pin.SequenceIndex
                                into g
                                                             select g;
            foreach (IGrouping<int, MeasPin> pinGroup in pinGroups)
            {
                int seqIndex = pinGroup.Key;
                bool isIdiff = pinGroup.ToList().Exists(s => s.MeasType == "MeasIdiff");
                bool isVdiff = pinGroup.ToList().Exists(s => s.MeasType == "MeasVdiff");
                if (isIdiff)
                {
                    pattern.MeasPins.RemoveAll(s => s.SequenceIndex == seqIndex && s.MeasType != "MeasIdiff");
                }
                if (isVdiff)
                {
                    pattern.MeasPins.RemoveAll(s => s.SequenceIndex == seqIndex && !s.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase));
                }
            }
            List<MeasPin> measCPins = pattern.MeasPins.FindAll(s => s.MeasType == MeasType.MeasC || s.MeasType == MeasType.MeasLimit);
            pattern.MeasPins.RemoveAll(measCPins.Contains);
            pattern.MeasPins = pattern.MeasPins.OrderBy(x => x.SequenceIndex).ToList();
            pattern.MeasPins.AddRange(measCPins);
        }

        private static List<MeasPin> ProcessSpecialMeasItems(List<MeasPin> allPins, bool isVdiff2, bool isVDiff, bool isUseTestPlan, bool isFlowUse)
        {
            var resultPins = new List<MeasPin>();
            if (allPins.Count == 0)
            {
                return resultPins;
            }

            int seqCount = allPins.Max(p => p.SequenceIndex);
            for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
            {
                var seqPins = allPins.Where(p => p.SequenceIndex == sequenceIndex).ToList();
                isVDiff = seqPins.Exists(p => p.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase));
                var inSeqPins = new List<MeasPin>();
                if (isVDiff && !isUseTestPlan)
                {//1. P,N 2. Vdiff 3. Vocm

                    foreach (MeasPin seqPin in seqPins)
                    {
                        if (!seqPin.MeasType.Equals(MeasType.MeasVocm, StringComparison.OrdinalIgnoreCase))
                        {
                            inSeqPins.Add(seqPin);
                        }
                        else if (isFlowUse)
                        {
                            inSeqPins.Add(seqPin);

                        }

                    }
                }
                else if (isVdiff2 && !isUseTestPlan)
                {
                    List<string> combinedDiff2Pins = DifferentialService.GroupDiffPairs(seqPins.Select(p => p.PinName).ToList());
                    foreach (string pin in combinedDiff2Pins)
                    {
                        var newPin = new MeasPin(pin, "MeasVdiff2")
                        {
                            SequenceIndex = sequenceIndex
                        };
                        inSeqPins.Add(newPin);
                    }
                }
                else if (seqPins.Count > 0)
                {
                    inSeqPins.AddRange(seqPins);
                }
                resultPins.AddRange(inSeqPins);
            }
            return resultPins;
        }

        private static bool JudgeSkipForDotNet(HardIpPattern pattern)
        {
            Dictionary<string, string> ignoreFunctions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Parse_IDS_Mapping_Table", "IDSCurrent" }, { "IDS_eFuse_Write", "IDSCurrent" }
            };
            if (ignoreFunctions.TryGetValue(pattern.FunctionName, out string ignoreFunction))
            {
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(ignoreFunction, "hardip");
                if (function.IsFound)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
