using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Basic;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using LogLib.Utility;

using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.BinNumber;
using TestPlanLib.BinNumberLegacy;
using TestPlanLib.Singleton;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility
{
    public class SearchInfo
    {
        private static readonly Regex _regex = new Regex("^((dd_)|(cz_)|(pp_)|(mn_)|(ht_)).*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool IsHardipIdsSheet(string sheetName)
        {
            return HardIpInputManager.RegexIds.IsMatch(sheetName);
        }

        public static bool IsHardipRtosSheet(string sheetName)
        {
            return HardIpInputManager.RegexRtos.IsMatch(sheetName);
        }
        public static bool IsSpiRomSheet(string sheetName)
        {
            return HardIpInputManager.RegexSpirom.IsMatch(sheetName);
        }


        public static List<HardIpInfo> GetHardIpInfos(List<string> patternNames, HardIpInfos hardIpInfos)
        {
            List<HardIpInfo> result = new List<HardIpInfo>();
            foreach (string patternName in patternNames)
            {
                if (hardIpInfos.TryGetValue(patternName, out List<HardIpInfo> items))
                {
                    result.AddRange(items);
                }
            }
            return result;
        }

        public static HardIpInfo ModDuplicateRegName(HardIpInfo hardIpInfo)
        {
            List<string> allreg = hardIpInfo.SendBitName.Split('+').ToList();
            List<string> assignment = hardIpInfo.DigSrcAssignment.Split(';').ToList();
            var assignfunc = new List<string>();
            var newAssignment = new List<string>();
            var trimreg = new List<string>();
            try
            {
                foreach (string assign in assignment)
                {
                    if (string.IsNullOrEmpty(assign))
                    {
                        continue;
                    }

                    if (!assign.Contains(':'))
                    {
                        trimreg.Add(assign);
                        continue;
                    }
                    if (assignfunc.Contains(assign.Split(':')[0]))
                    {
                        newAssignment.Add(assign.Split(':')[0] + '_' + assign.Split(':')[1] + ':' + assign.Split(':')[1]);
                        if (allreg.Contains(assign.Split(':')[0]))
                        {
                            int indx = allreg.FindLastIndex(x => x.Equals(assign.Split(':')[0]));
                            allreg[indx] = assign.Split(':')[0] + '_' + assign.Split(':')[1];
                        }
                    }
                    else
                    {
                        assignfunc.Add(assign.Split(':')[0]);
                        newAssignment.Add(assign);
                    }
                }
                if (trimreg.Count > 0)
                {
                    newAssignment.AddRange(trimreg);
                }

                if (hardIpInfo.TrimRegName.Count > 0)
                {
                    int index = 0;
                    foreach (string trimfuse in hardIpInfo.TrimFuseName.Split(','))
                    {
                        if (index < hardIpInfo.TrimRegName.Count)
                        {
                            newAssignment.Add(string.Format("{0}:{1}", hardIpInfo.TrimRegName[index], trimfuse));
                        }

                        index++;
                    }

                }
                hardIpInfo.SendBitName = string.Join("+", allreg);
                hardIpInfo.DigSrcAssignment = string.Join(";", newAssignment);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return hardIpInfo;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetTestLimitPerMeasType(HardIpPattern pattern)
        {
            var dicLimintPerpin = new Dictionary<string, string>();
            HardIpInfo patInfo = HardIpService.GetHardIpInfo(pattern);
            int seqCount = patInfo == null || patInfo.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : patInfo.SeqInfo.Count;

            #region Gets TestLimitPerPin_VFI for vif
            if (Regex.IsMatch(pattern.FunctionName, VbtFunctionLibShared.VifName, RegexOptions.IgnoreCase))
            {
                dicLimintPerpin.Add("V", "F");
                dicLimintPerpin.Add("F", "F");
                dicLimintPerpin.Add("I", "F");

                #region exists Sequence in patInfo
                for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                {
                    SetVifLimitPerMeasTypeBySequence(pattern, sequenceIndex, dicLimintPerpin);
                }
                return dicLimintPerpin;
                #endregion
            }

            #endregion

            #region Gets TestLimitPerPin_VIR for vir

            if (Regex.IsMatch(pattern.FunctionName, VbtFunctionLibShared.VirName, RegexOptions.IgnoreCase))
            {
                dicLimintPerpin.Add("V", "F");
                dicLimintPerpin.Add("I", "F");
                dicLimintPerpin.Add("R", "F");

                #region exists Sequence in patInfo

                for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                {
                    SetVirLimitPerMeasTypeBySequence(pattern, sequenceIndex, dicLimintPerpin);
                }
                return dicLimintPerpin;

                #endregion
            }

            #endregion

            return dicLimintPerpin;
        }

        private static void SetVifLimitPerMeasTypeBySequence(HardIpPattern pattern, int sequenceIndex, Dictionary<string, string> dicLimintPerpin)
        {
            List<string> pinlist = new List<string>();
            string measType = "";
            foreach (MeasPin measPin in pattern.MeasPins)
            {
                if (measPin.SequenceIndex == sequenceIndex && !Regex.IsMatch(measPin.MeasType, "Calc|Limit|MeasC", RegexOptions.IgnoreCase))
                {
                    pinlist.Add(measPin.PinName);
                    measType = measPin.MeasType;
                }

            }
            if (pinlist.Count == 1 && !pinlist[0].Contains(",") && !pinlist[0].Contains("::"))
            {
                return;
            }

            if (measType.Equals("MeasV", StringComparison.OrdinalIgnoreCase) || measType.Equals("MeasVdiff", StringComparison.OrdinalIgnoreCase))
            {
                dicLimintPerpin["V"] = "T";
            }
            if (measType.Equals("MeasI", StringComparison.OrdinalIgnoreCase) || measType.Equals("MeasI2", StringComparison.OrdinalIgnoreCase) || measType.Equals("MeasIdiff", StringComparison.OrdinalIgnoreCase))
            {
                dicLimintPerpin["I"] = "T";
            }
            if (measType.Equals("MeasF", StringComparison.OrdinalIgnoreCase) || measType.Equals("MeasFdiff", StringComparison.OrdinalIgnoreCase))
            {
                dicLimintPerpin["F"] = "T";
            }
        }

        private static void SetVirLimitPerMeasTypeBySequence(HardIpPattern pattern, int sequenceIndex, Dictionary<string, string> dicLimintPerpin)
        {
            List<string> pinlist = new List<string>();
            string measType = "";
            foreach (MeasPin measPin in pattern.MeasPins)
            {
                if (measPin.SequenceIndex == sequenceIndex && !Regex.IsMatch(measPin.MeasType, "Calc|Limit|MeasC", RegexOptions.IgnoreCase))
                {
                    pinlist.Add(measPin.PinName);
                    measType = measPin.MeasType;
                }
            }
            if (pinlist.Count == 1 && !pinlist[0].Contains(",") && !pinlist[0].Contains("::"))
            {
                return;
            }

            if (measType.Equals("MeasV", StringComparison.OrdinalIgnoreCase) || measType.Equals("MeasE", StringComparison.OrdinalIgnoreCase))
            {
                dicLimintPerpin["V"] = "T";
            }
            if (measType.Equals("MeasI", StringComparison.OrdinalIgnoreCase))
            {
                dicLimintPerpin["I"] = "T";
            }
            if (measType.Equals("MeasR1", StringComparison.OrdinalIgnoreCase))
            {
                dicLimintPerpin["R"] = "T";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static bool GetFlagSingleLimit(HardIpPattern pattern, string labelVoltage)
        {
            if (Regex.IsMatch(pattern.CalcEqn, "TX_Level", RegexOptions.IgnoreCase))
            {
                return false;
            }

            if (Regex.IsMatch(pattern.MiscInfo, "Flag_Singlelimit"))
            {
                foreach (string flagStr in pattern.MiscInfo.Split(';'))
                {
                    if (Regex.IsMatch(pattern.MiscInfo, "Flag_Singlelimit"))
                    {
                        try
                        {
                            string flagResult = flagStr.Split(':')[1].Trim();
                            return bool.Parse(flagResult);
                        }
                        catch (Exception)
                        {
                            //if parse error, continue judge wihtout crash
                        }
                    }
                }
            }

            HardIpInfo patInfo = HardIpService.GetHardIpInfo(pattern);
            int seqCount = patInfo == null || patInfo.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : patInfo.SeqInfo.Count;

            #region exists Sequence in patInfo
            for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
            {
                var sequenceLimits = pattern.MeasPins.Where(p => p.SequenceIndex == sequenceIndex && !Regex.IsMatch(p.MeasType, "Calc|Limit|MeasC", RegexOptions.IgnoreCase)).ToList();
                if (sequenceLimits.Exists(p => p.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase)))
                {
                    return false;
                }

                if (sequenceLimits.Exists(p => p.MeasType.Equals(MeasType.MeasVdiff2, StringComparison.OrdinalIgnoreCase)))
                {
                    return false;
                }

                if (sequenceLimits.GroupBy(p => p.PinName).Count() > 1)
                {
                    return false;
                }
            }
            #endregion

            return true;
        }

        public static string GetCpuflag(HardIpInfo info, HardIpPattern pattern)
        {
            if (string.IsNullOrEmpty(pattern.Pattern.GetLastPayload()) ||
                pattern.Pattern.GetLastPayload().Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase) ||
                HardIpConstData.RegInsInPatt.IsMatch(pattern.Pattern.GetLastPayload()) ||
                !string.IsNullOrEmpty(pattern.BlockType))
            {
                return "false";
            }

            if (string.IsNullOrEmpty(info.CallSubrs))
            {
                return "false";
            }
            else
            {
                return "true";
            }
        }

        public static string GetStoreName(HardIpPattern pattern, HardIpInputData hardIpInputData, ref string storeNameOri, string testSequence = "")
        {
            var storeNameList = new List<string>();
            var pinsInSeq = pattern.MeasPins.GroupBy(p => p.SequenceIndex).ToDictionary(p => p.Key, p => p.ToList());

            foreach (KeyValuePair<int, List<MeasPin>> pins in pinsInSeq)
            {
                if (pins.Key < 1)
                {
                    continue;
                }

                var storeNameSeqList = new List<string>();
                var oriNames = new List<string>();
                foreach (MeasPin pin in pins.Value)
                {
                    if (pin.PinName.Contains("CP=") && !string.IsNullOrEmpty(pin.CusStr))
                    {
                        storeNameSeqList.Add("CP=" + pin.CusStr);
                        oriNames.Add(pin.CusStr);
                    }
                    else if (pin.PinName.Contains("FT=") && !string.IsNullOrEmpty(pin.CusStr))
                    {
                        storeNameSeqList.Add("FT=" + pin.CusStr);
                        oriNames.Add(pin.CusStr);
                    }
                    else
                    {
                        storeNameSeqList.Add(pin.CusStr);
                        oriNames.Add(pin.CusStr);
                    }
                }

                if (oriNames.Distinct(StringComparer.OrdinalIgnoreCase).Count(p => !string.IsNullOrEmpty(p)) > 1)//storeNameSeqList.Count)
                {
                    storeNameList.Add(string.Join(":", storeNameSeqList));
                }
                else
                {
                    storeNameList.Add(oriNames.FirstOrDefault(p => !string.IsNullOrEmpty(p)));
                }
            }
            string storeNameText = storeNameList.Any(p => !string.IsNullOrEmpty(p)) ? string.Join("+", storeNameList) : "";

            storeNameOri = storeNameText;
            if (storeNameText.Length >= 6000)
            {
                var registerAssignItem = new HardIpRegAssign();
                string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
                registerAssignItem.SubBlockName = blockName + "_" + CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName);
                if (Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^dd_", RegexOptions.IgnoreCase))
                {
                    registerAssignItem.SubBlockName += "_DD";
                }

                registerAssignItem.Type = RegisterAssignType.Meas_StoreName;
                if (!hardIpInputData.HardIpRegAssigns.Exists(x => x.SubBlockName.Equals(registerAssignItem.SubBlockName, StringComparison.CurrentCultureIgnoreCase) && x.Type == registerAssignItem.Type))
                {
                    var regAssginList = new List<List<string>>();
                    if (string.IsNullOrEmpty(testSequence))
                    {
                        List<string> storeNames = storeNameText.Split('+').ToList();
                        for (int index = 0; index < storeNames.Count; index++)
                        {
                            string storeName = storeNames[index];
                            List<string> data = new List<string> { storeName };
                            regAssginList.Add(data);
                        }
                    }
                    else
                    {
                        List<string> testSequenceArr = testSequence.Split(',').ToList();
                        List<string> storeNames = storeNameText.Split('+').ToList();
                        if (testSequenceArr.Count == storeNames.Count)
                        {
                            for (int index = 0; index < storeNames.Count; index++)
                            {
                                string storeName = storeNames[index];
                                List<string> data = new List<string> { testSequenceArr.ElementAt(index), storeName };
                                regAssginList.Add(data);
                            }
                        }
                    }

                    registerAssignItem.RegAssignList = regAssginList;
                    hardIpInputData.HardIpRegAssigns.Add(registerAssignItem);
                }
                storeNameText = $"Reg_assign:{registerAssignItem.SubBlockName}";
            }
            return storeNameText;

        }

        /// <summary>
        /// if just one pinGroup in a each meas squence. then the pingrop should not be decomposed
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="measPin"></param>
        /// <param name="dicTestLimitPerPin"></param>
        /// <returns></returns>
        public static bool NeedDecompGroups(HardIpPattern pattern, MeasPin measPin, Dictionary<string, string> dicTestLimitPerPin)
        {
            if (/*pattern.FunctionName == VbtFunctionLibShared.VdiffFunc || */measPin.MeasType.Equals("MeasR2"))
            {
                return true;
            }

            if (dicTestLimitPerPin.ContainsKey(measPin.MeasType[4].ToString()) &&
                dicTestLimitPerPin[measPin.MeasType[4].ToString()] == "F")
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// return MeasC pin name from CapPinName in patInfo, the default value is "JTAG_TDO"
        /// </summary>
        /// <param name="pattern"></param>
        /// <returns>MeasC pin name</returns>
        public static string GetMeasCPins(HardIpPattern pattern, HardIpInfo hardIpInfo)
        {
            if (hardIpInfo.CapBitStr != "")
            {
                if (hardIpInfo.CapPinName != "")
                {
                    return hardIpInfo.CapPinName.Split('|').First();
                }

                return "JTAG_TDO";
            }
            return "";
        }

        public static bool IsMeasPinInForcePin(string forcePinGroup, string measPinName, List<string> allForceConditionPins = null)
        {
            if (forcePinGroup.Equals(measPinName))
            {
                return true;
            }

            if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(forcePinGroup))
            {
                List<string> forcePinGroupList = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(forcePinGroup);
                if (forcePinGroupList.Contains(measPinName))
                {
                    return true;
                }

                var measPins = new List<string> { measPinName };
                if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(measPinName))
                {
                    measPins = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(measPinName);

                }
                if (measPins.All(forcePinGroupList.Contains) || (allForceConditionPins != null && measPins.Count > allForceConditionPins.Count ? measPins.All(allForceConditionPins.Contains) : allForceConditionPins.All(p => measPins.Contains(p))))
                {
                    return true;
                }

                return false;
            }
            else
            {
                var measPins = new List<string> { measPinName };
                if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(measPinName))
                {
                    measPins = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(measPinName);
                }
                if (allForceConditionPins != null && measPins.Count > allForceConditionPins.Count ? measPins.All(allForceConditionPins.Contains) : allForceConditionPins.All(p => measPins.Contains(p)))
                {
                    return true;
                }

                return false;
            }
        }

        public static string GetSrcPin(HardIpInfo patInfo)
        {
            if ((patInfo.SendBit != "0" && !string.IsNullOrEmpty(patInfo.SendBit)) || patInfo.SendBitStr != "")
            {
                string srcPin = string.IsNullOrEmpty(patInfo.SendPinName) ? "JTAG_TDI" : patInfo.SendPinName.Split('|').First();
                return srcPin;
            }
            return "";
        }

        /// <summary>
        /// Get pin type accoding to PinMap. Default type is "I/O"
        /// </summary>
        /// <param name="pinName"></param>
        /// <returns>Pin Type</returns>
        public static string GetPinType(string pinName)
        {
            if (Regex.IsMatch(pinName, "^VDD", RegexOptions.IgnoreCase))
            {
                return "power";
            }

            string name = pinName;
            if (name.Contains("::"))
            {
                name = pinName.Split(':')[0];
            }
            else if (name.Contains(":"))
            {
                name = pinName.Split(':')[1];
            }
            else if (name.Contains("="))
            {
                name = pinName.Split('=')[1];
            }

            if (TestProgram.IgxlWorkBk.PinMapPair.Value.TryGetPin(name, out Pin pin))
            {
                return pin.PinType;
            }

            if (TestProgram.IgxlWorkBk.PinMapPair.Value.TryGetGroup(name, out PinGroup group))
            {
                return group.PinType;
            }

            return "I/O";
        }

        /// <summary>
        /// Calculate MeasC count of each pattern
        /// </summary>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static int MeasC_Count(HardIpPattern pattern)
        {
            return pattern.MeasPins.Where(pin => pin.MeasType == "MeasC").Sum(pin => pin.PinCount);
        }

        public static void GetPlanCurrentRange(List<MeasPin> measPins, List<MeasPin> patInfoPins, bool isRepeatLimit)
        {
            if (measPins.Exists(p => p.MeasType.Equals(MeasType.WiSrc) || p.MeasType.Equals(MeasType.WiMeas)))
            {
                patInfoPins = measPins;
                return;
            }

            foreach (MeasPin planMeasPin in measPins)
            {
                planMeasPin.IsUsedPin = false;
            }
            var vocmPins = new List<MeasPin>();
            var patInfoPinsGroups = patInfoPins.GroupBy(p => p.SequenceIndex).ToDictionary(p => p.Key, p => p.ToList());
            var resultPins = new List<MeasPin>();

            foreach (KeyValuePair<int, List<MeasPin>> patInfoPinsGroup in patInfoPinsGroups)
            {
                var seqPins = new List<MeasPin>();
                foreach (MeasPin patInfoPin in patInfoPinsGroup.Value)
                {
                    // prevent the same name in patinfo , that can't get the second measure pin
                    var groupPins = new List<string>();

                    int cnt = measPins.Where(p => !p.IsUsedPin && p.MeasType.Equals(patInfoPin.MeasType.Split('>')[0], StringComparison.OrdinalIgnoreCase)).Count(x => ContainsPin(x.PinName, patInfoPin.PinName, ref groupPins) && (x.VisitedTime > 0 || isRepeatLimit));
                    // prevent the same name in patinfo , that can't get the second measure pin
                    bool isFoundPin = false;
                    MatchPlanMeasPin(measPins, patInfoPin, seqPins, cnt, isRepeatLimit, ref groupPins, ref isFoundPin);
                    #region
                    if (patInfoPin.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase))
                    {
                        CollectVocmPin(measPins, patInfoPin, vocmPins, cnt, isRepeatLimit, ref groupPins);
                    }
                    #endregion

                    if (!isFoundPin && seqPins.All(p => !ContainsPin(p.PinName, patInfoPin.PinName, ref groupPins)))
                    {
                        var oriPatInfoPin = new MeasPin(patInfoPin.OriPinName, patInfoPin.MeasType.Split('>')[0])
                        {
                            SequenceIndex = patInfoPin.SequenceIndex
                        };
                        if (!seqPins.Any(p => p.PinName.Equals(oriPatInfoPin.PinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            seqPins.Add(oriPatInfoPin);
                        }
                    }
                }

                resultPins.AddRange(seqPins);
            }

            resultPins.AddRange(vocmPins);
            patInfoPins.Clear();
            patInfoPins.AddRange(resultPins);
            ResetVisiteTime(measPins);
        }

        private static void MatchPlanMeasPin(List<MeasPin> measPins, MeasPin patInfoPin, List<MeasPin> seqPins, int cnt, bool isRepeatLimit, ref List<string> groupPins, ref bool isFoundPin)
        {
            for (int i = 0; i < measPins.Count; i++)
            {
                MeasPin planMeasPin = measPins[i];
                if (cnt > 0)
                {
                    if (ContainsPin(planMeasPin.PinName, patInfoPin.PinName, ref groupPins) &&
                        patInfoPin.MeasType.Split('>')[0].Equals(planMeasPin.MeasType, StringComparison.OrdinalIgnoreCase) &&
                        (planMeasPin.VisitedTime > 0 || isRepeatLimit) && !planMeasPin.IsUsedPin)
                    {
                        isFoundPin = true;
                        if (seqPins.Exists(p =>
                            p.PinName.Equals(planMeasPin.PinName, StringComparison.OrdinalIgnoreCase) &&
                            p.SequenceIndex != planMeasPin.SequenceIndex))
                        {
                            continue;
                        }

                        //patInfoPin.Copy(planMeasPin); //.CurrentRange = planMeasPin.CurrentRange;
                        planMeasPin.SequenceIndex = patInfoPin.SequenceIndex;
                        planMeasPin.IsUsedPin = true;
                        patInfoPin.PinName = planMeasPin.PinName;
                        seqPins.Add(planMeasPin);
                        if (planMeasPin.MeasType == "MeasVdiff2")
                        {
                            planMeasPin.VisitedTime -= 2;
                        }
                        else
                        {
                            //planMeasPin.VisitedTime--;
                            planMeasPin.VisitedTime = 0;
                        }
                        break;
                    }

                }
                else
                {
                    if (ContainsPin(planMeasPin.PinName, patInfoPin.PinName, ref groupPins) &&
                        patInfoPin.MeasType.Split('>')[0].Equals(planMeasPin.MeasType, StringComparison.OrdinalIgnoreCase) &&
                        (planMeasPin.VisitedTime > 0 || isRepeatLimit))
                    {
                        isFoundPin = true;
                        if (
                            seqPins.Exists(
                                p =>
                                    p.PinName.Equals(planMeasPin.PinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            continue;
                        }

                        planMeasPin.SequenceIndex = patInfoPin.SequenceIndex;
                        planMeasPin.IsUsedPin = true;
                        patInfoPin.PinName = planMeasPin.PinName;
                        seqPins.Add(patInfoPin);
                        if (planMeasPin.MeasType == "MeasVdiff2")
                        {
                            planMeasPin.VisitedTime -= 2;
                        }
                        else
                        {
                            planMeasPin.VisitedTime = 0;
                        }
                        break;
                    }

                }
            }
        }

        private static void CollectVocmPin(List<MeasPin> measPins, MeasPin patInfoPin, List<MeasPin> vocmPins, int cnt, bool isRepeatLimit, ref List<string> groupPins)
        {
            var vocmPin = new MeasPin();
            foreach (MeasPin planMeasPin in measPins)
            {
                if (cnt > 0)
                {
                    if (ContainsPin(planMeasPin.PinName, patInfoPin.PinName, ref groupPins) && planMeasPin.MeasType.Equals(MeasType.MeasVocm, StringComparison.OrdinalIgnoreCase))
                    {
                        vocmPin.Copy(planMeasPin);
                        planMeasPin.IsUsedPin = true;
                        vocmPin.PinName = patInfoPin.PinName;
                        vocmPin.SequenceIndex = patInfoPin.SequenceIndex;
                        vocmPins.Add(vocmPin);
                        break;
                    }
                }
                else
                {
                    if (ContainsPin(planMeasPin.PinName, patInfoPin.PinName, ref groupPins) && patInfoPin.MeasType.Equals(planMeasPin.MeasType, StringComparison.OrdinalIgnoreCase) && (planMeasPin.VisitedTime > 0 || isRepeatLimit))
                    {
                        vocmPin.Copy(planMeasPin);
                        planMeasPin.IsUsedPin = true;
                        vocmPin.SequenceIndex = patInfoPin.SequenceIndex;
                        vocmPin.PinName = patInfoPin.PinName;
                        vocmPins.Add(vocmPin);
                        if (planMeasPin.MeasType == "MeasVdiff2")
                        {
                            planMeasPin.VisitedTime -= 2;
                        }
                        else
                        {
                            planMeasPin.VisitedTime = 0;
                        }
                        break;
                    }
                }

            }
        }

        internal static bool ContainsPin(string planPinName, string patInfoPinName, ref List<string> pinList)
        {
            if (planPinName == patInfoPinName)
            {
                return true;
            }

            string newPlanPinName = planPinName;
            string newPatInfoPinName = patInfoPinName;
            if (!patInfoPinName.Contains("::") && patInfoPinName.Contains(":"))
            {
                newPatInfoPinName = patInfoPinName.Split(':')[1];
            }

            if (patInfoPinName.Contains("::") && planPinName.Contains("::"))
            {
                var infoDiffpins = Regex.Split(patInfoPinName, "::", RegexOptions.IgnoreCase).OrderBy(p => p).ToList();
                var planDiffpins = Regex.Split(planPinName, "::", RegexOptions.IgnoreCase).OrderBy(p => p).ToList();
                if (infoDiffpins.SequenceEqual(planDiffpins, StringComparer.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            if (patInfoPinName.Contains("::") && planPinName.Contains(","))
            {
                var infoDiffpins = Regex.Split(patInfoPinName, "::", RegexOptions.IgnoreCase).OrderBy(p => p).ToList();
                var planDiffpins = Regex.Split(planPinName, ",", RegexOptions.IgnoreCase).OrderBy(p => p).ToList();
                if (infoDiffpins.SequenceEqual(planDiffpins, StringComparer.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            if (newPlanPinName.Equals(newPatInfoPinName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(planPinName))
            {
                List<string> pinGroupList = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(planPinName);
                string checkPin = newPatInfoPinName.Contains('=') ? newPatInfoPinName.Split('=')[1] : newPatInfoPinName;
                if (pinGroupList.Contains(checkPin))
                {
                    pinList = pinGroupList;
                    return true;
                }
            }
            return false;
        }

        private static void ResetVisiteTime(List<MeasPin> measPins)
        {
            foreach (MeasPin planMeasPin in measPins)
            {
                planMeasPin.VisitedTime = GetPinCount(planMeasPin.PinName);
            }
        }

        internal static int GetPinCount(string pinNames)
        {
            int count = 0;
            foreach (string pin in pinNames.Split(','))
            {
                if (pin == "")
                {
                    continue;
                }

                if (pin.Contains("::"))
                {
                    count += 2;
                }
                else if (TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(pin))
                {
                    List<string> pinGroupList = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(pin);
                    count += pinGroupList.Count;
                }
                else
                {
                    count++;
                }
            }
            return count;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static string GetCusStrDigCapData(HardIpPattern pattern, HardIpInfo info, bool isNeedMerge = true, bool isCsUsing = false)
        {
            #region Get DsscOut from pattern info && check if Disable_MeasC_Split
            if ((Regex.IsMatch(pattern.MiscInfo, HardIpConstData.KeepDsscOut, RegexOptions.IgnoreCase) || LocalSpecs.Options.Device == EnumDevice.RF) && info.CapBitStr != "" && info.CapBitName != "" && info.DsscOut != "")
            {
                RecomputeDsscOut(info);
            }
            List<string> dsscout = Regex.Replace(info.DsscOut.Trim(','), "DSSC_OUT:", "DSSC_OUT,", RegexOptions.IgnoreCase).Split(',').ToList();
            #endregion

            #region When set "ignore_patt_comment", user can change bits of measC by testPlan

            bool ignorePatInfo = Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatInfo, RegexOptions.IgnoreCase);
            bool ignorePatMeasC = Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatMeasC, RegexOptions.IgnoreCase);
            bool isMrrEnable = Regex.IsMatch(pattern.MiscInfo, HardIpConstData.StoreDigAll, RegexOptions.IgnoreCase);
            string capData = "";
            ApplyIgnorePatOverrides(pattern, info, ignorePatInfo, ignorePatMeasC, ref dsscout, ref capData);

            #endregion

            #region Replace DSSC OUT by testName/CusStr in testPlan

            int capPinIndex = 1;
            foreach (MeasPin pin in pattern.MeasPins)
            {
                if (pin.MeasType != MeasType.MeasC)
                {
                    continue;
                }

                if (dsscout.Count > capPinIndex)
                {
                    string[] strArr = dsscout[capPinIndex].Split(':');
                    if (strArr.Length == 2)
                    {
                        if (!string.IsNullOrEmpty(pin.TestName))
                        {
                            strArr[1] = pin.TestName;
                        }

                        dsscout[capPinIndex] = string.Join(":", strArr.ToList());
                        if (!string.IsNullOrEmpty(pin.CusStr))
                        {
                            dsscout[capPinIndex] = dsscout[capPinIndex] + ":" + pin.CusStr;
                        }
                    }
                }
                capPinIndex++;
            }
            capData = string.Join(",", dsscout);
            #endregion

            string trimOrCapName = GetDigCapNameByMiscInfo(pattern.MiscInfo);
            if (trimOrCapName != "")
            {
                capData = (trimOrCapName + "&" + capData).Trim('&');
            }

            if (isCsUsing)
            {
                return capData;
            }

            return capData == "" ? "" : isMrrEnable ? HardIpConstData.StoreDigAll + "&" + capData : capData;
        }

        private static void RecomputeDsscOut(HardIpInfo info)
        {
            string[] bitStrArray = info.CapBitStr.Split(new[] { '|', '+' }, StringSplitOptions.RemoveEmptyEntries);
            string[] bitNameArray = info.CapBitName.Split(new[] { '|', '+' }, StringSplitOptions.RemoveEmptyEntries);
            string oriDsscOut = Regex.Replace(info.DsscOut, @"DSSC_OUT.\s*", "", RegexOptions.IgnoreCase).Trim(',');
            string[] dsscOutArray = oriDsscOut.Split(',');
            var newDsscOut = new List<string>();
            int index = 0;
            int nameIndex = 0;
            foreach (string bitStr in bitStrArray)
            {
                int width = Convert.ToInt32(bitStr.Split('_')[1]);
                int dsscOutWidth = Convert.ToInt32(dsscOutArray[index].Split(':')[0]);
                if (width == dsscOutWidth)
                {
                    newDsscOut.Add(width + ":" + bitNameArray[nameIndex]);
                }
                else
                {
                    while (width > dsscOutWidth)
                    {
                        index++;
                        int nextdsscOutWidth = Convert.ToInt32(dsscOutArray[index].Split(':')[0]);
                        dsscOutWidth += nextdsscOutWidth;
                    }
                    if (width == dsscOutWidth)
                    {
                        newDsscOut.Add(width + ":" + bitNameArray[nameIndex]);
                    }
                }
                index++;
                nameIndex++;
            }
            info.DsscOut = "DSSC_OUT," + string.Join(",", newDsscOut);
        }

        private static void ApplyIgnorePatOverrides(HardIpPattern pattern, HardIpInfo info, bool ignorePatInfo, bool ignorePatMeasC, ref List<string> dsscout, ref string capData)
        {
            if (ignorePatMeasC)
            {
                if (pattern.CapBitsInTp() == info.CapBit && info.CapBit != 0)
                {
                    dsscout = new List<string> { "DSSC_OUT" };
                    foreach (
                        MeasPin pinMeasC in
                            pattern.MeasPins.Where(
                                p => p.MeasType.Equals(MeasType.MeasC, StringComparison.OrdinalIgnoreCase)))
                    {
                        dsscout.Add($"{pinMeasC.CapBit}:{pinMeasC.CusStr}");
                    }
                    capData = string.Join(",", dsscout);
                }
            }
            else if (ignorePatInfo)
            {
                if (pattern.CapBitsInTp() == info.CapBit && info.CapBit != 0)
                //Total bits in TestPlan and info are the same
                {
                    bool isCapBitsMatch = IsCapBitsMatch(pattern, info);
                    if (dsscout.Count - 1 == MeasC_Count(pattern) && MeasC_Count(pattern) != 0 &&
                        !isCapBitsMatch)
                    {
                        // the seguence of patinfo and T/P are not the same
                        dsscout.RemoveRange(0, dsscout.Count);
                        dsscout.Add("DSSC_OUT");
                        foreach (MeasPin measCpin in pattern.MeasPins.FindAll(s => s.MeasType == MeasType.MeasC))
                        {
                            dsscout.Add(measCpin.CapBit + ":MeasC_" + dsscout.Count);
                        }
                    }
                }
            }
        }

        public static bool IsCapBitsMatch(HardIpPattern patternItem, HardIpInfo patInfo)
        {
            bool capBitsFlag = true;

            string[] bitStrArray = patInfo.CapBitStr.Split('+');
            int bitStrArrayCnt = patInfo.CapBitStr == "" ? 0 : patInfo.CapBitStr.Split('+').Length;
            var measClist = patternItem.MeasPins.Where(x => x.MeasType == MeasType.MeasC).ToList();
            if (bitStrArrayCnt == measClist.Count)
            {
                if (bitStrArrayCnt != 0)
                {
                    int cnt = 0;
                    foreach (string bitStr in bitStrArray)
                    {
                        int width = 0;
                        int capBit = 0;
                        if (int.TryParse(bitStr.Split('_')[1], out int value1))
                        {
                            width = value1;
                        }

                        if (int.TryParse(measClist[cnt].CapBit, out int value2))
                        {
                            capBit = value2;
                        }

                        if (width != capBit)
                        {
                            capBitsFlag = false;
                            break;
                        }
                        cnt++;
                    }
                }
            }
            else
            {
                capBitsFlag = false;
            }
            return capBitsFlag;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="miscInfo"></param>
        /// <returns></returns>
        /// 
        public static List<string> GetTrimStoreNameByMiscInfo(string miscInfo)
        {
            var trimList = new List<string>();
            string digTrimNameReg = HardIpConstData.TrimStoreName + @"\w*:(?<name>.*)";
            string digTrimCodeNameReg = HardIpConstData.TrimCodeStoreName + @"\w*:(?<name>.*)";
            string capStrDigCapData = HardIpConstData.CusStrDigCapData + @"\w*:(?<name>.*)";
            string dictStoreCodeName = HardIpConstData.DictStoreCodeName + @"\w*:(?<name>.*)";
            string trimDictionaryStoreName = HardIpConstData.TrimDictionaryStoreName + @"\w*:(?<name>.*)";
            string digCapDataCustomString = HardIpConstData.DigCapDataCustomString + @"\w*:(?<name>.*)";
            //Dict_Store_Code_Name
            foreach (string item in miscInfo.Split(';'))
            {
                if (Regex.IsMatch(item, digTrimNameReg, RegexOptions.IgnoreCase))
                {
                    trimList.AddRange(Regex.Split(Regex.Match(item, digTrimNameReg, RegexOptions.IgnoreCase).Groups["name"].ToString(), "[,+&]", RegexOptions.IgnoreCase));
                }
                if (Regex.IsMatch(item, digTrimCodeNameReg, RegexOptions.IgnoreCase))
                {
                    trimList.AddRange(Regex.Split(Regex.Match(item, digTrimCodeNameReg, RegexOptions.IgnoreCase).Groups["name"].ToString(), "[,+&]", RegexOptions.IgnoreCase));
                }
                if (Regex.IsMatch(item, capStrDigCapData, RegexOptions.IgnoreCase))
                {
                    trimList.AddRange(Regex.Split(Regex.Match(item, capStrDigCapData, RegexOptions.IgnoreCase).Groups["name"].ToString(), "[,+&]", RegexOptions.IgnoreCase));
                }
                if (Regex.IsMatch(item, dictStoreCodeName, RegexOptions.IgnoreCase))
                {
                    trimList.AddRange(Regex.Split(Regex.Match(item, dictStoreCodeName, RegexOptions.IgnoreCase).Groups["name"].ToString(), "[,+&]", RegexOptions.IgnoreCase));
                }
                if (Regex.IsMatch(item, trimDictionaryStoreName, RegexOptions.IgnoreCase))
                {
                    trimList.AddRange(Regex.Split(Regex.Match(item, trimDictionaryStoreName, RegexOptions.IgnoreCase).Groups["name"].ToString(), "[,+&]", RegexOptions.IgnoreCase));
                }
                if (Regex.IsMatch(item, digCapDataCustomString, RegexOptions.IgnoreCase))
                {
                    trimList.AddRange(Regex.Split(Regex.Match(item, digCapDataCustomString, RegexOptions.IgnoreCase).Groups["name"].ToString(), "[,+&]", RegexOptions.IgnoreCase));
                }
            }
            return trimList;
        }

        public static string GetDigCapNameByMiscInfo(string miscInfo)
        {
            string digCapNameReg = HardIpConstData.DigCapName + ":(?<name>.*)";
            foreach (string item in miscInfo.Split(';'))
            {
                if (Regex.IsMatch(item, digCapNameReg, RegexOptions.IgnoreCase))
                {
                    return Regex.Match(item, digCapNameReg, RegexOptions.IgnoreCase).Groups["name"].ToString();
                }
            }
            return "";
        }

        /// <summary>
        /// Get specified VBT module from Misc Info
        /// Change to use vbtnameMapping on 2016/6/27
        /// </summary>
        /// <param name="hardIpInputData"></param>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static string GetVbtNameByPattern(HardIpInputData hardIpInputData, HardIpPattern pattern)
        {
            #region Judge TrimTarget as trim function
            if (!string.IsNullOrEmpty(pattern.WirelessData.TrimTarget))
            {
                pattern.VbtTypes.Add(PlanType.Trim);
            }
            #endregion

            #region if exist specified function Name => use it!
            foreach (string info in pattern.MiscInfo.Split(';'))
            {
                if (!Regex.IsMatch(info.Trim(), "^(" + HardIpConstData.Vbt + @")\s*\:\s*\w+", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                KeyValuePair<string, string> mappingPair = hardIpInputData.ConfigData.VbtNameMapping.FirstOrDefault(a => a.Key.Equals(info.Trim(), StringComparison.OrdinalIgnoreCase));
                if (mappingPair.Value != null)
                {
                    Function mappingVbt = TestProgram.VbtFunctionLib.GetFunctionByName(mappingPair.Value, "hardip");
                    if (mappingVbt.IsFound)
                    {
                        return mappingPair.Value;
                    }
                }

                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(info.Split(':')[1], "hardip");
                if (!string.IsNullOrEmpty(function.FunctionName))
                {
                    if (!function.IsFound)
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_MissVbtModule_02,
                            pattern.SheetName,
                            pattern.RowNum,
                            0,
                            [info.Split(':')[1]]
                        );
                    }

                    return info.Split(':')[1];
                }
            }
            #endregion

            if (IsHardipIdsSheet(pattern.SheetName))
            {
                return VbtFunctionLibShared.Ids;
            }

            if (pattern.SheetName.StartsWith(NeededSheets.PrefixLcd, StringComparison.OrdinalIgnoreCase))
            {
                HardIpInfo hardip = HardIpService.GetHardIpInfo(pattern.Pattern.GetLastPayload());
                if (!string.IsNullOrEmpty(pattern.WirelessData.TrimTarget))
                {
                    return VbtFunctionLibShared.LcdTrim;
                }

                return VbtFunctionLibShared.LcdMeas;
            }

            if (LocalSpecs.Options.Device == EnumDevice.LCD && pattern.SheetName.StartsWith("DCTEST_", StringComparison.OrdinalIgnoreCase))
            {
                return VbtFunctionLibShared.LcdMeas;
            }

            if (pattern.MeasPins.Exists(p => p.MeasType.Equals(MeasType.WiSrc, StringComparison.OrdinalIgnoreCase) || p.MeasType.Equals(MeasType.WiMeas, StringComparison.OrdinalIgnoreCase)))
            {
                if (!string.IsNullOrEmpty(pattern.WirelessData.TrimTarget))
                {
                    HardIpService.GetHardIpInfo(pattern.Pattern.GetLastPayload());
                    return VbtFunctionLibShared.RfTrim;
                }

                return VbtFunctionLibShared.RfFunc;
            }

            if (!string.IsNullOrEmpty(pattern.WirelessData.TrimTarget))
            {
                HardIpInfo info = HardIpService.GetHardIpInfo(pattern.Pattern.GetLastPayload());
                if (info.TrimRegName.Count > 1)
                {
                    return VbtFunctionLibShared.DvdcTrim3D;
                }

                return VbtFunctionLibShared.DvdcTrim;
            }
            return "";
        }

        /// <summary>
        /// Set the env in Flow sheet
        /// </summary>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static string GetEnvFromPattern(HardIpPattern pattern)
        {
            if (LocalSpecs.IsPatternValidate)
            {
                return "";
            }

            string patternName = pattern.Pattern.GetLastPayload();
            if (patternName.Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase) ||
                HardIpConstData.RegInsInPatt.IsMatch(pattern.Pattern.RealPatternName) ||
                HardIpConstData.RegOpcodeInPatt.IsMatch(pattern.Pattern.RealPatternName))
            {
                return "";
            }

            if ((pattern.SheetName.Equals("Wireless_ARF", StringComparison.OrdinalIgnoreCase) || pattern.SheetName.Equals("HardIP_ARF", StringComparison.OrdinalIgnoreCase)) && LocalSpecs.Options.Device == EnumDevice.RF)
            {
                return "";
            }

            bool missInPatternList = GetStatusInPatternList(AcTSetCategoryMapSingleton.Instance().PatternList, patternName, out string status) ||
                                     pattern.FunctionName.Equals(VbtFunctionLibShared.RtosRunScenarioT, StringComparison.OrdinalIgnoreCase);
            if (!missInPatternList)
            {
                return status;
            }

            return "";
        }

        public static bool GetStatusInPatternList(Dictionary<string, PatternData> patternList, string patternName, out string status)
        {
            status = "";
            if (!patternList.ContainsKey(patternName))
            {
                status = "MissPattInPattList";
            }
            else if (patternList[patternName].TimeSetVersion.ToLower() == "na")
            {
                status = "MissTimesetInPattList";
            }
            else if (patternList[patternName].FileVersion == "NA")
            {
                status = "MissFileVersionInPattList";
            }

            if (status != "")
            {
                return false;
            }

            return true;
        }

        public static string GetTtrEnable(string ttrStr, string voltage, List<string> joblist = null)
        {
            ttrStr = ttrStr.Replace(" ", "");
            List<string> allJobs = joblist?.Any() == true ? joblist : LocalSpecs.AllJobsHardIp.ToList();
            foreach (string ttr in ttrStr.Split(';'))
            {
                if (ttr == "")
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(voltage) && ttr.Contains(voltage))
                {
                    // [JobName] or [NV/LV/HV]
                    if (!ttr.Contains(":"))
                    {
                        // [NV/LV/HV]
                        if (ttr.EndsWith("V"))
                        {
                            if (ttr.Split(',').Any(ttrvol => ttrvol.Equals(voltage, StringComparison.OrdinalIgnoreCase)))
                            {
                                allJobs.Clear();
                            }
                        }
                        // [JobName]
                        else
                        {
                            allJobs = allJobs.Where(a => !Regex.IsMatch(a, ttr)).ToList();
                        }
                    }
                    // [JobName]:[NV/LV/HV]
                    else
                    {
                        string jobName = ttr.Split(':')[0];
                        allJobs = allJobs.Where(a => !Regex.IsMatch(a, jobName)).ToList();
                    }
                }
                else
                {
                    // eg: CP1 means that NV,HV,LV should be removed from CP1
                    allJobs = allJobs.Where(a => !Regex.IsMatch(a, ttr)).ToList();
                }
            }
            return string.Join(",", allJobs);
        }

        /// <summary>
        /// Reset relay setting
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public static string ReverseRelaySetting(string setting)
        {
            string relayOn = Regex.Match(setting, "(?<relayon>RelayOff[^R]+).*").Groups["relayon"].ToString().Replace("Off", "On");
            string relayOff = Regex.Match(setting, "(?<relayoff>RelayOn[^R]+).*").Groups["relayoff"].ToString().Replace("On", "Off");
            return (relayOn + "_" + relayOff).Trim('_');
        }

        /// <summary>
        /// eg. RelayOn_K1_K2_RelayOff_K3 ==>  argOn:K1,K2  argOff:K3
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="argOn"></param>
        /// <param name="argOff"></param>
        public static void GetRelayArgs(string setting, out string argOn, out string argOff)
        {
            string outOn = "";
            string outOff = "";
            string relayOn = Regex.Match(setting, "RelayOn:(?<relayon>[^R]+).*").Groups["relayon"].ToString().Trim('_');
            string relayOff = Regex.Match(setting, "RelayOff:(?<relayoff>[^R]+).*").Groups["relayoff"].ToString().Trim('_');
            if (relayOn != "")
            {
                List<string> onList = relayOn.Split('&').ToList();
                outOn = string.Join(",", onList);
            }
            if (relayOff != "")
            {
                List<string> offList = relayOff.Split('&').ToList();
                outOff = string.Join(",", offList);
            }
            //If both on/off exist, relayOn is "K1_K2_", so argOn may be "K1,K2,"
            argOn = outOn.Trim(',');
            argOff = outOff;
        }

        public static bool IsForceType(HardIpPattern pattern, string type)
        {
            return pattern.MeasPins.Any(x => x.ForceConditions.Any(y => y.ForcePins.Any(z => z.ForceType.StartsWith(type))));
        }

        public static bool IsMeasVdiff2(HardIpPattern pattern)
        {
            HardIpInfo patInfo = HardIpService.GetHardIpInfo(pattern);
            return pattern.MeasPins.Any(pin => pin.MeasType.ToLower() == "measvdiff2") || (patInfo != null && patInfo.MeasVdiff2PinList.Count > 0);
        }

        public static bool IsMeasVdiff(HardIpPattern pattern)
        {
            HardIpInfo patInfo = HardIpService.GetHardIpInfo(pattern);
            return pattern.MeasPins.Any(pin =>
                pin.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase)) ||
                (patInfo != null && patInfo.MeasVdiffPinList.Count > 0);
        }

        public static bool IsRepeatLimit(string miscInfo)
        {
            return miscInfo.Split(';').ToList().Exists(s => Regex.IsMatch(s, HardIpConstData.RepeatLimit, RegexOptions.IgnoreCase));
        }

        public static int GetRepeatLimitCount(string miscInfo)
        {
            int loop = 0;
            foreach (string assign in miscInfo.Split(';'))
            {
                string[] asignArr = assign.Split(':');
                if (asignArr.Length == 2 && Regex.IsMatch(asignArr[0], HardIpConstData.RepeatLimit, RegexOptions.IgnoreCase))
                {
                    int.TryParse(asignArr[1].Trim(), out loop);
                }
            }
            return loop;
        }

        public static List<string> GetOpcode(HardIpPattern pattern, string opType)
        {
            string miscInfo = pattern.MiscInfo;
            var opcodeListA = new List<string>(); //After the pattern
            var opcodeListB = new List<string>(); //Before the pattern
            foreach (string info in miscInfo.Split(';'))
            {
                if (!Regex.IsMatch(info, @"opcode\s*\:\w+\s*(\:.*)?", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                string code = Regex.Match(info, @"opcode\s*\:(?<code>(\w+))\s*(\:.*)?", RegexOptions.IgnoreCase).Groups["code"].ToString();
                string name = Regex.Match(info, @"opcode\s*\:\w+\s*(\:\s*(?<name>(.*)))?", RegexOptions.IgnoreCase).Groups["name"].ToString();
                if (opType == "B" && !(code.ToLower() == "endif" || code.ToLower() == "next") && pattern.FlowControlFlag != 1)
                {
                    opcodeListB.Add(code + ":" + name);
                }

                if (opType == "A" && (code.ToLower() == "endif" || code.ToLower() == "next") && pattern.FlowControlFlag != 0)
                {
                    opcodeListA.Add(code + ":" + name);
                }
            }
            if (opType == "A")
            {
                return opcodeListA;
            }

            return opcodeListB;
        }

        public static string GetFlagNoBinStr(string miscInfo, string voltage)
        {
            foreach (string item in miscInfo.Split(';'))
            {
                if (Regex.IsMatch(item, HardIpConstData.NoBinUseLimit + "|" + HardIpConstData.NoBin, RegexOptions.IgnoreCase))
                {
                    List<string> noBinVoltage = item.Split(':').ToList();
                    if (noBinVoltage.Count > 1 && Regex.IsMatch(noBinVoltage[1], voltage))
                    {
                        return "No";
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// If there is more MeasC in patInfo, create MeasC pins for test plan to match the count in patInfo
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="patInfo"></param>
        public static void InitialMeasC(ref HardIpPattern pattern, HardIpInfo patInfo)
        {
            if (patInfo.DsscOut != "")
            {
                int planCount = MeasC_Count(pattern);
                int infoCount = patInfo.DsscOut.Trim(',').Split(',').Length - 1;
                if (Regex.IsMatch(pattern.MiscInfo, HardIpConstData.KeepDsscOut, RegexOptions.IgnoreCase))
                {
                    if (planCount == Regex.Split(patInfo.CapBitName, @"\+|\|").Length)
                    {
                        return;
                    }
                }
                if (infoCount > planCount && !Regex.IsMatch(pattern.MiscInfo, HardIpConstData.IgnorePatMeasC, RegexOptions.IgnoreCase) &&
                    LocalSpecs.Options.Device != EnumDevice.LCD && LocalSpecs.Options.Device != EnumDevice.RF)
                {
                    string pinName = patInfo.CapPinName == "" ? "JTAGTDO" : patInfo.CapPinName;
                    if (pattern.Pattern.IsMultiTimeDomain() && pattern.Pattern.IsFullMtdPattern())
                    {
                        for (int patindex = 0; patindex < patInfo.MultiDsscOut.Split('|').Length; patindex++)
                        {
                            string patternName = pattern.Pattern.TestPlanPatternName.Split('#', '+')[patindex].Trim();

                            int singlePlanCount = pattern.MeasPins.Where(pin =>
                                pin.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase) &&
                                pin.PatternIndex.Equals(patindex) &&
                                pin.MeasType == "MeasC").Sum(pin => pin.PinCount);
                            int singleInfoCount = patInfo.MultiDsscOut.Split('|')[patindex].Trim(',').Split(',').Length - 1;

                            if (singleInfoCount > singlePlanCount)
                            {
                                for (int j = 0; j < singleInfoCount - singlePlanCount; j++)
                                {
                                    var pin = new MeasPin(pinName, "MeasC")
                                    {
                                        PinCount = 1,
                                        PatternName = patternName,
                                        PatternIndex = patindex
                                    };
                                    pattern.MeasPins.Add(pin);
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < infoCount - planCount; i++)
                        {
                            var pin = new MeasPin(pinName, "MeasC")
                            {
                                PinCount = 1
                            };
                            pattern.MeasPins.Add(pin);
                        }
                    }
                }
            }
        }

        public static BinNumResult GetRtosBin(HardIpPattern pattern)
        {
            string blockName = Regex.Replace(CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName), "wireless_|lcd_", "", RegexOptions.IgnoreCase);
            string subBlockName = CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName);
            string binName = string.IsNullOrEmpty(pattern.Failflag) ? $"{HardIpConstData.BinFlowFlag}_{pattern}_{subBlockName}" : pattern.Failflag.Replace("F_", "Bin_");
            string module = "RTOS";
            string[] binNameSegments = binName.Split('_');
            string category1 = binNameSegments.Length > 1 ? binNameSegments[1] : "";
            string category2 = binNameSegments.Length > 2 ? binNameSegments[2] : "";
            return BinNumberSingleton.Instance.GetBinInfo(module, category1, category2);
        }

        public static BinNumResult GetHardIpBin(HardIpPattern pattern)
        {
            string module = pattern.SheetName.Split('_')[0];
            string category1 = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            string category2 = pattern.SubBlock.Replace(" ", "").Replace("_", "").Replace("-", "");
            return BinNumberSingleton.Instance.GetBinInfo(module, category1, category2);
        }

        /// <summary>
        /// Use scgh to search BinNumber setting in config
        /// </summary>
        /// <param name="pattern"></param>
        /// <returns></returns>
        public static SoftBinRangeData GetHardIpBinRangeItem(HardIpPattern pattern, ProdCharSheetRow scghRow)
        {
            const string index5 = "HardIP_others";
            SoftBinRangeData binRange;
            var para = new BinNumDefPara(EnumBinNumDefBlock.HardIp, index5);
            if (scghRow != null)
            {
                string index1 = scghRow.Block + "-" + scghRow.Mode + "-" + scghRow.Item;
                string index2 = scghRow.Block + "-" + scghRow.Mode;
                string index3 = scghRow.Block;
                string index4 = scghRow.Block.Replace(".", "").Replace("p", "");
                if (pattern.SheetName.ContainsIgnoreCase("DCTEST_IDS"))
                {
                    if (pattern.Pattern.GetAliasPatternList().Any(x => x.Contains("_nan")))
                    {
                        para.Condition = "IDS_NAND";
                        binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
                        if (binRange != null)
                        {
                            return binRange;
                        }
                    }
                    else if (pattern.Pattern.GetAliasPatternList().Any(x => x.Contains("_spi")))
                    {
                        para.Condition = "IDS_SPI";
                        binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
                        if (binRange != null)
                        {
                            return binRange;
                        }
                    }
                }

                para.Condition = index1;
                binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
                if (binRange != null)
                {
                    return binRange;
                }

                para.Condition = index2;
                binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
                if (binRange != null)
                {
                    return binRange;
                }

                para.Condition = index3;
                binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
                if (binRange != null)
                {
                    return binRange;
                }

                para.Condition = index4;
                binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
                if (binRange != null)
                {
                    return binRange;
                }
            }

            binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
            if (binRange != null)
            {
                return binRange;
            }

            para.Condition = Regex.Replace(pattern.SheetName, "HardIP_|Wireless_|LCD_", "", RegexOptions.IgnoreCase).Replace("_", "");
            binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
            if (binRange != null)
            {
                return binRange;
            }

            para.Condition = index5;
            binRange = BinNumberSingletonLegacy.Instance().SearchSoftBinRangeData(para);
            if (binRange != null)
            {
                return binRange;
            }

            //const string errorMessage = "Missing bin number setting";
            ErrorReportManager.AddError(
                HardIpErrorType.E_MissingBinNum_01,
                pattern.SheetName,
                pattern.RowNum,
                0,
                []
            );
            return new SoftBinRangeData();
        }

        /// <summary>
        /// Eg: "Calc:Algorithm_A;CalcParameter:rd0,rd1", return "Alg::Algorithm_A(rd0,rd1)"
        /// </summary>
        /// <param name="miscInfo"></param>
        /// <param name="testName"></param>
        /// <returns></returns>
        public static string GetCalculation(string miscInfo, string testName = "")
        {
            string calculation = "";
            List<string> calculationList = new List<string>();
            if (miscInfo.IndexOf(HardIpConstData.Calc + ":", StringComparison.OrdinalIgnoreCase) != -1)
            {
                string vol = "";
                string algrothim = "";
                string paras = "";
                bool flag = false;
                bool setflag = false;
                foreach (string item in miscInfo.Split(';'))
                {
                    Match getCalc = Regex.Match(item, "^(?<vol>(NV|LV|HV)@)?Calc:", RegexOptions.IgnoreCase);
                    if (getCalc.Success && !flag)
                    {
                        vol = getCalc.Groups["vol"].Value;
                        algrothim = item.Split(':')[1];
                        flag = true;
                    }

                    if ((item.StartsWith(HardIpConstData.CalcParameter + ":", StringComparison.OrdinalIgnoreCase) ||
                        item.StartsWith(HardIpConstData.CalcParameter + "s:", StringComparison.OrdinalIgnoreCase) ||
                        item.StartsWith(HardIpConstData.NvCalcParameter + ":", StringComparison.OrdinalIgnoreCase) ||
                        item.StartsWith(HardIpConstData.NvCalcParameter + "s:", StringComparison.OrdinalIgnoreCase) ||
                        item.StartsWith(HardIpConstData.LvCalcParameter + ":", StringComparison.OrdinalIgnoreCase) ||
                        item.StartsWith(HardIpConstData.LvCalcParameter + "s:", StringComparison.OrdinalIgnoreCase) ||
                        item.StartsWith(HardIpConstData.HvCalcParameter + ":", StringComparison.OrdinalIgnoreCase) ||
                        item.StartsWith(HardIpConstData.HvCalcParameter + "s:", StringComparison.OrdinalIgnoreCase)) && flag)

                    {
                        paras = item.Substring(item.IndexOf(":", StringComparison.Ordinal) + 1, item.Length - item.IndexOf(":", StringComparison.Ordinal) - 1);
                        flag = false;
                        setflag = true;
                    }

                    if (setflag && algrothim != "" && paras != "")
                    {
                        calculationList.Add(vol + "Alg:" + testName + ":" + algrothim + "(" + paras + ")");
                        setflag = false;
                    }
                }
            }
            if (calculationList.Count > 0)
            {
                calculation = string.Join(";", calculationList);
            }

            return calculation;
        }

        public static string TrimSpace(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }

            char[] buffer = input.ToCharArray();
            int idx = 0;

            for (int i = 0; i < buffer.Length; i++)
            {
                char c = buffer[i];
                if (c != ' ' && c != '\n' && c != '\r' && c != '\t')
                {
                    buffer[idx++] = c;
                }
            }

            if (idx == input.Length)
            {
                return input;
            }

            return new string(buffer, 0, idx);
        }

        public static List<string> GetMeasStrByPlan(HardIpPattern pattern)
        {
            var measlst = new List<string>();
            for (int seqIndex = 1; seqIndex <= pattern.TestPlanSequences.Count; seqIndex++)
            {
                MeasPin firstOrDefault = pattern.MeasPins.FirstOrDefault(s => s.SequenceIndex == seqIndex && s.MeasType != "MeasC" && s.PinName != "FakePin"
                                                                              && s.MeasType != "");
                if (firstOrDefault != null)
                {
                    if (firstOrDefault.MeasType.Equals(MeasType.WiMeas))
                    {
                        measlst.Add("A");
                    }
                    else if (firstOrDefault.MeasType.Equals(MeasType.WiSrc))
                    {
                        measlst.Add("G");
                    }
                    else if (firstOrDefault.MeasType.Equals(MeasType.MeasWait))
                    {
                        measlst.Add("W");
                    }
                    else
                    {
                        measlst.Add(firstOrDefault.MeasType.Replace("Meas", ""));
                    }
                }

                else if (pattern.TestPlanSequences[seqIndex - 1].ForceCondition.Count > 0)
                {
                    measlst.Add("N");
                }
            }
            if (measlst.Count == 1 && measlst[0] == "N")
            {
                measlst.Clear();
            }

            return measlst;
        }

        public static string GenDiffGroupName(string diffPinName, bool isNeedGenPinGroup)
        {
            string groupName = "";
            if (!diffPinName.Contains("::"))
            {
                return diffPinName;
            }

            if (!isNeedGenPinGroup)
            {
                return string.Join(",", Regex.Split(diffPinName, "::"));
            }
            if (TestProgram.IgxlWorkBk.PinMapPair.Value != null)
            {
                string[] pair = diffPinName.Split(new[] { "::" }, StringSplitOptions.None);
                groupName = TestProgram.IgxlWorkBk.PinMapPair.Value.GetDiffGroupName(pair);
            }

            if (groupName == "")
            {
                DifferentialService.DiffPinPosAndNeg(diffPinName, out string _, out string _, out groupName);
            }

            if (groupName == "")
            {
                return diffPinName;
            }

            return groupName;
        }

        public static string GetDigDataWidth(string sendBitStr, string defaultValue = "")
        {
            string reg = @"[a-zA-Z]+\d+_(?<num>(\d+))";
            MatchCollection matches = Regex.Matches(sendBitStr, reg, RegexOptions.IgnoreCase);
            if (matches.Count == 0)
            {
                return string.Empty;
            }

            string width = matches[0].Groups["num"].ToString();
            if (matches.Cast<Match>().Any(match => !match.Groups["num"].ToString().Equals(width)))
            {
                return defaultValue;
            }
            return width;
        }

        //To do Delete
        public static string GetPpmuPin(HardIpPattern pattern, HardIpInfo info)
        {
            string measPPmu = "";
            var seqMeas = new List<string>();
            int seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
            if (seqCount <= 0)
            {
                return measPPmu;
            }

            for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
            {
                var measPinList = pattern.MeasPins.Where(a => a.SequenceIndex == sequenceIndex && a.MeasType != "MeasC" && a.MeasType != MeasType.MeasLimit && a.MeasType != "").ToList();
                List<string> internalMeas = measPinList.Select(measPin => measPin.PinName).ToList();
                seqMeas.Add(string.Join(",", internalMeas));
            }
            measPPmu = string.Join("+", seqMeas).Replace("::", ","); //Convert P1_P::P1_N to P1_P,P1_N(differential pins)

            return measPPmu;
        }

        public static string GetForceV(HardIpPattern pattern, HardIpInfo info, string voltage = "", bool isCsUsing = false)
        {
            if (IsForceType(pattern, "V"))
            {
                string forceV = "";
                int seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
                List<string> seqList = new List<string>();
                if (seqCount > 0)
                {
                    for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                    {
                        var measPinList = pattern.MeasPins.Where(a => a.SequenceIndex == sequenceIndex && a.MeasType != MeasType.MeasC).ToList();
                        bool isAllMeasR2 = IsAllTheSameType(measPinList, MeasType.MeasR2);
                        IsAllTheSameType(measPinList, MeasType.MeasI);
                        IsAllTheSameType(measPinList, MeasType.MeasV);
                        string forceDelimiter = isAllMeasR2 ? "&" : ",";

                        List<string> fieldList = new List<string>();
                        foreach (MeasPin measPin in measPinList)
                        {
                            if (measPin.ForceConditions.Count > 0)
                            {

                                var allForcePins = measPin.ForceConditions.SelectMany(forcecondition =>
                                    forcecondition.ForcePins.SelectMany(forcepin => TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(forcepin.PinName))).ToList();
                                foreach (ForceCondition condition in measPin.ForceConditions)
                                {
                                    List<string> forceVPerPin = BuildForceVPerPin(measPin, condition, allForcePins, voltage);
                                    if (measPin.MeasType == MeasType.MeasVdiff2 && isCsUsing)
                                    {
                                        fieldList.Add(string.Join(forceDelimiter, GetForcePairValueDotNet(forceVPerPin)));
                                    }
                                    else
                                    {
                                        fieldList.Add(string.Join(forceDelimiter, forceVPerPin));
                                    }
                                }
                            }
                            else
                            {
                                fieldList.Add("");
                            }
                        }
                        if (fieldList.Distinct().Count() == 1)
                        {
                            seqList.Add(fieldList[0]);
                        }
                        else
                        {
                            if (isAllMeasR2 && isCsUsing)
                            {
                                seqList.Add(string.Join(",", GetForcePairValueDotNet(fieldList)));
                            }
                            else
                            {
                                seqList.Add(string.Join(",", fieldList));
                            }
                        }
                    }
                    forceV = string.Join("|", seqList);

                    return RemoveDummyForceV(forceV, @",|\+|\|");
                }
            }
            return "";
        }

        private static List<string> BuildForceVPerPin(MeasPin measPin, ForceCondition condition, List<string> allForcePins, string voltage)
        {
            List<string> forceVPerPin = new List<string>();
            foreach (ForcePin forcePin in condition.ForcePins)
            {
                string measPinName = measPin.PinName.Split(':').Length == 2 ? measPin.PinName.Split(':')[1] : measPin.PinName;
                if (measPin.MeasType == MeasType.MeasVdiff2)
                {
                    if (measPinName.Contains("::"))
                    {
                        if (!(IsMeasPinInForcePin(forcePin.PinName, measPinName.Split(':')[0], allForcePins) ||
                            IsMeasPinInForcePin(forcePin.PinName, measPinName.Split(':')[2], allForcePins))
                            || forcePin.Type == ForceConditionType.Others)
                        {
                            continue;
                        }
                    }
                }
                else
                {
                    if (!IsMeasPinInForcePin(forcePin.PinName, measPinName, allForcePins) || forcePin.Type == ForceConditionType.Others)
                    {
                        continue;
                    }
                }
                if (forcePin.ForceLabelVoltages.Count > 0 && voltage != null &&
                !forcePin.ForceLabelVoltages.Contains(voltage))
                {
                    continue;
                }

                if (forcePin.ForceType == "V")
                {
                    if (forcePin.ForceJob == "")
                    {
                        forceVPerPin.Add(DataConvertor.ConvertForceValueToGlbSpec(forcePin));
                    }
                    else
                    {
                        forceVPerPin.Add(forcePin.ForceJob + ":" + DataConvertor.ConvertForceValueToGlbSpec(forcePin));
                    }
                }
            }
            return forceVPerPin;
        }

        public static string GetForceI(HardIpPattern pattern, HardIpInfo info, string voltage = "")
        {
            if (IsForceType(pattern, "I"))
            {
                IsAllTheSameType(pattern.MeasPins, "measi");
                string forceDelimiter = ",";
                string forceI = "";
                int seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
                List<string> seqList = new List<string>();
                if (seqCount > 0)
                {
                    for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                    {
                        var measPinList = pattern.MeasPins.Where(a => a.SequenceIndex == sequenceIndex && a.MeasType != "MeasC").ToList();
                        List<string> fieldList = new List<string>();
                        foreach (MeasPin measPin in measPinList)
                        {
                            List<string> forceIPerPin = new List<string>();
                            if (measPin.ForceConditions.Count > 0)
                            {
                                var allForcePins = measPin.ForceConditions.SelectMany(forcecondition =>
                                    forcecondition.ForcePins.SelectMany(forcepin => TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(forcepin.PinName))).ToList();
                                foreach (ForceCondition condition in measPin.ForceConditions)
                                {
                                    foreach (ForcePin forcePin in condition.ForcePins)
                                    {
                                        string measPinName = measPin.PinName.Split(':').Length == 2 ? measPin.PinName.Split(':')[1] : measPin.PinName;
                                        if (!IsMeasPinInForcePin(forcePin.PinName, measPinName, allForcePins) || forcePin.Type == ForceConditionType.Others)
                                        {
                                            continue;
                                        }

                                        if (forcePin.ForceLabelVoltages.Count > 0 && voltage != null &&
                                            !forcePin.ForceLabelVoltages.Contains(voltage))
                                        {
                                            continue;
                                        }

                                        if (forcePin.ForceType == "I")
                                        {
                                            if (forcePin.ForceJob == "")
                                            {
                                                forceIPerPin.Add(DataConvertor.ConvertForceValueToGlbSpec(forcePin));
                                            }
                                            else
                                            {
                                                forceIPerPin.Add(forcePin.ForceJob + ":" +
                                                                 DataConvertor.ConvertForceValueToGlbSpec(forcePin));
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                forceIPerPin.Add("");
                            }

                            fieldList.Add(string.Join(forceDelimiter, forceIPerPin.Distinct().Count() == 1 ? forceIPerPin.Distinct() : forceIPerPin));
                        }
                        if (fieldList.Distinct().Count() == 1)
                        {
                            seqList.Add(fieldList[0]);
                        }
                        else
                        {
                            seqList.Add(string.Join(",", fieldList));
                        }
                    }
                    forceI = string.Join("|", seqList);
                    return RemoveDummyForceV(forceI, @",|\+|\&|\|");
                }
            }
            return "";
        }

        internal static string RemoveDummyForceV(string forceV, string regexString)
        {
            if (forceV != "")
            {
                List<string> values = Regex.Split(forceV, regexString).ToList();
                int count = values.Distinct().ToList().Count;
                if (count == 1)
                {
                    return values[0];
                }
            }
            return forceV;
        }

        private static string RemoveDummy(string value, string regexString)
        {
            if (value != "")
            {
                List<string> valueList = Regex.Split(value, regexString, RegexOptions.IgnoreCase).ToList();
                if (valueList.All(x => x == ""))
                {
                    return "";
                }
            }
            return value;
        }

        public static string GetIRange(HardIpPattern pattern, HardIpInfo info, string voltage)
        {
            string currentRange = "";
            List<string> sequenceList = new List<string>();
            int seqCount = info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
            if (seqCount > 0)
            {
                for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                {
                    List<string> currentRangeList = new List<string>();

                    var measPinList = pattern.MeasPins.Where(a => a.SequenceIndex == sequenceIndex && a.MeasType != "MeasC"
                        && !a.PinName.StartsWith("FT", StringComparison.OrdinalIgnoreCase)).ToList();

                    string lastftNetName = "";
                    string range = "";
                    foreach (MeasPin measPin in measPinList)
                    {
                        bool isMeasR = measPinList.Exists(x => x.MeasType == MeasType.MeasR1 || x.MeasType == MeasType.MeasR2);
                        ForceConditionType mode = isMeasR ? GetForceMode(measPinList) : ForceConditionType.Normal;
                        if (mode == ForceConditionType.FiMode)
                        {
                            currentRangeList.Add("");
                            continue;
                        }
                        string cpNetName = "";
                        string ftNetName = "";
                        TestPlanStatic.PowerMergeSheet.PowerMerge.GetCpFtNetName(measPin.PinName.Split('=').Last(),
                            ref cpNetName, ref ftNetName);
                        if (!string.IsNullOrEmpty(lastftNetName) && ftNetName == lastftNetName &&
                            measPin.PinName.Contains('='))
                        {
                            range = new MeasPinCurrentRangeCalculator(measPin).GetCurrentRangeByVoltageCp(voltage);
                        }
                        else
                        {
                            range = new MeasPinCurrentRangeCalculator(measPin).GetCurrentRangeByVoltage(voltage);
                        }

                        range = range.Replace("999.999", "");

                        currentRangeList.Add(range);
                        if (measPin.PinName.Contains("::"))
                        {
                            currentRangeList.Add(range);
                        }

                        lastftNetName = ftNetName;
                    }
                    bool isAllEqual = currentRangeList.Distinct().Count() == 1;
                    //sequenceList.Add(currentRangeList.Any(x => x != "") ? string.Join(",", currentRangeList) : "");
                    sequenceList.Add(currentRangeList.Any(x => x != "") ? isAllEqual ? string.Join(",", currentRangeList.Distinct()) : string.Join(",", currentRangeList) : "");
                }
                currentRange = string.Join("+", sequenceList);
            }

            if (pattern.MeasPins.Exists(s => Regex.IsMatch(s.MeasType, "^(measi|measidiff|(MeasR[1|2]))", RegexOptions.IgnoreCase)))
            {
                currentRange = GetIRangeByJob(currentRange);
            }

            if (currentRange == "0")
            {
            }
            return RemoveDummy(currentRange, @"\+|,");
        }

        public static string GetPrePat(HardIpPattern pattern, string voltage = "")
        {
            List<string> interposePrePats = new List<string>();
            List<string> forceConditionPrePats = new List<string>();
            List<string> sweepCodePrePats = new List<string>();
            bool isSweepFirst = GetTheForceConditionOrder(pattern.ForceCondition.ForceCondition);
            if (pattern.ForceConditionList.Count > 0)
            {
                BuildForceConditionPrePats(pattern, voltage, forceConditionPrePats);
            }

            #region insert SweepVoltage and SweepCodes information

            if (pattern.SweepVoltage.Count >= 1)
            {
                BuildSweepCodePrePats(pattern, voltage, sweepCodePrePats);
            }

            if (isSweepFirst)
            {
                interposePrePats.AddRange(sweepCodePrePats);
                interposePrePats.AddRange(forceConditionPrePats);
            }
            else
            {
                interposePrePats.AddRange(forceConditionPrePats);
                interposePrePats.AddRange(sweepCodePrePats);
            }
            string.Join(";", interposePrePats);

            #endregion

            #region MiscInfo: USL, LSL

            string miscStr = pattern.MiscInfo;
            foreach (string param in miscStr.Split(';'))
            {
                if (!Regex.IsMatch(param, "^USL:" + "|" + "^LSL:", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                interposePrePats.Add(param.Trim());
            }

            #endregion

            if (pattern.ForceCondition.IsVtShmoo && pattern.IsVtShmooUsed)
            {
                foreach (string forceCondition in pattern.ForceCondition.ForceCondition.Split(';'))
                {
                    Match match = HardIpConstData.RegVtShmoo.Match(forceCondition);
                    if (match.Success)
                    {
                        string shmooStr = match.Groups["ShmooStr"].Value;
                        string[] strArray = shmooStr.Split(':');

                        if (strArray.Length > 2)
                        {
                            interposePrePats.Add($"{strArray[0]}:{strArray[1]}:SHMOO_GLB");
                        }
                    }
                }
                pattern.IsVtShmooUsed = false;
            }

            #region Add shmoo global for force
            BuildShmooGlobalForForce(pattern, voltage, interposePrePats);
            #endregion

            return string.Join(";", interposePrePats);
        }

        private static void BuildForceConditionPrePats(HardIpPattern pattern, string voltage, List<string> forceConditionPrePats)
        {
            foreach (ForceCondition condition in pattern.ForceConditionList)
            {
                foreach (ForcePin pin in condition.ForcePins)
                {
                    if (pin.ForceLabelVoltages.Count > 0 && voltage != null &&
                        !pin.ForceLabelVoltages.Contains(voltage))
                    {
                        continue;
                    }

                    if (pin.Type == ForceConditionType.Normal)
                    {
                        if (pin.ForceJob == "")
                        {
                            forceConditionPrePats.Add(pin.PinName + ":" + pin.ForceType + ":" + DataConvertor.ConvertForceValueToGlbSpec(pin, "", true) + (pin.IsRestore ? "" : ":NoRestore"));
                        }
                        else
                        {
                            forceConditionPrePats.Add(pin.PinName + ":" + pin.ForceType + ":" + DataConvertor.ConvertForceValueToGlbSpec(pin, "", true) + ":" + pin.ForceJob);
                        }
                    }
                    if (pin.Type == ForceConditionType.Others)
                    {
                        forceConditionPrePats.Add(pin.PinName + ":" + pin.ForceValue);
                    }
                }
            }
        }

        private static void BuildSweepCodePrePats(HardIpPattern pattern, string voltage, List<string> sweepCodePrePats)
        {
            foreach (KeyValuePair<string, List<SweepVData>> sVitems in pattern.SweepVoltage)
            {
                string sweepVoltageType = sVitems.Key;
                foreach (SweepVData item in sVitems.Value.Where(p => !string.IsNullOrEmpty(p.Type)
                                                                     && (string.IsNullOrEmpty(p.InstanceVoltage)
                                                                     || voltage.Equals(p.InstanceVoltage, StringComparison.OrdinalIgnoreCase))))
                {
                    if (sVitems.Key == "NestSweep")
                    {
                        sweepVoltageType = !string.IsNullOrEmpty(item.ForceType) ? item.ForceType : "V";
                    }

                    if (!string.IsNullOrEmpty(item.Order))// for nestsweep, no need for loop source code idx
                    {
                        foreach (string pin in item.PinName.Split(','))
                        {
                            string one = $"{pin}:{sweepVoltageType}:sweepvoltage{item.Order}{item.CustomSetting}";
                            sweepCodePrePats.Add(one);
                        }
                    }
                    else
                    {
                        foreach (string pin in item.PinName.Split(','))
                        {
                            string one =
                                $"{pin}:{sweepVoltageType}:{item.Start}{item.Operand}[SrcCodeIndex{item.Axis}]*{item.Step}{item.CustomSetting}";
                            sweepCodePrePats.Add(one);
                        }
                    }
                }
            }
        }

        private static void BuildShmooGlobalForForce(HardIpPattern pattern, string voltage, List<string> interposePrePats)
        {
            foreach (string forceCondition in pattern.ForceCondition.ForceCondition.Split(';'))
            {
                string[] forceVoltageCondition = forceCondition.Split('@');
                if (forceVoltageCondition.Length > 1 && !string.IsNullOrEmpty(forceVoltageCondition[0]))
                {
                    if (!forceVoltageCondition[0].Equals(voltage, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }

                Match match = HardIpConstData.RegShmoo.Match(forceCondition);
                if (match.Success)
                {
                    string shmooStr = match.Groups["ShmooStr"].Value;
                    string[] strArr = shmooStr.Replace("::", "&&").Split(':'); // diff pin group workaround

                    if (strArr.Length < 3)
                    {
                        continue;
                    }

                    string firstWord = strArr[1].Split(',')[0];

                    string[] dollarArr = firstWord.Split('$');
                    if (dollarArr.Length == 2)
                    {
                        string type = dollarArr[0];
                        string shmooName = dollarArr[1];

                        if (!string.IsNullOrEmpty(type))
                        {
                            string pin = DataConvertor.ConvertToNetName(strArr[0], TestPlanStatic.PowerMergeSheet.PowerMerge);

                            string symbol = "_Shmoo_" + shmooName + "_Glb";
                            interposePrePats.Add(string.Format("{0}:{1}:{2}", pin, type, symbol));
                        }
                    }
                }
            }
        }

        private static bool GetTheForceConditionOrder(string forceCondition)
        {
            bool ret = false;
            if (string.IsNullOrEmpty(forceCondition))
            {
                return ret;
            }

            if (forceCondition.ContainsIgnoreCase("SWEEP") && forceCondition.ContainsIgnoreCase("DISCONNECT"))
            {
                return forceCondition.ToUpper().StartsWith("SWEEP");
            }

            return ret;
        }

        public static string GetPostMeas(HardIpPattern pattern, string testSequence = "", string voltage = "", string preMeasDelimiter = "|")
        {
            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            List<string> infoSequence = info.MeasSeqStr == "" ? GetMeasStrByPlan(pattern) : info.MeasSeqStr.Split(',').ToList();

            var postMeas = new List<string>();
            if (infoSequence.Count > 0)
            {
                for (int index = 1; index <= infoSequence.Count; index++)
                {
                    CollectPostMeasBySequence(pattern, index, voltage, postMeas);
                }
                return postMeas.All(string.IsNullOrEmpty) ? "" : string.Join(preMeasDelimiter, postMeas);
            }
            return "";
        }

        private static void CollectPostMeasBySequence(HardIpPattern pattern, int index, string voltage, List<string> postMeas)
        {
            var measPins = pattern.MeasPins.Where(a => a.SequenceIndex == index).ToList();
            var postMeasSeq = new List<string>();
            var allForcePins = new List<ForcePin>();
            //Collect all force pins in the same sequenceIndex
            if (!(pattern.FunctionName.Equals(VbtFunctionLibShared.VifName, StringComparison.OrdinalIgnoreCase) ||
                  pattern.FunctionName.Equals(VbtFunctionLibShared.VirName, StringComparison.OrdinalIgnoreCase) ||
                  pattern.FunctionName.Equals(VbtFunctionLibShared.LcdMeas, StringComparison.OrdinalIgnoreCase) ||
                  pattern.FunctionName.Equals(VbtFunctionLibShared.LcdTrim, StringComparison.OrdinalIgnoreCase) ||
                  pattern.FunctionName.Equals(VbtFunctionLibShared.DvdcTrim3D, StringComparison.OrdinalIgnoreCase) ||
            pattern.FunctionName.Equals(VbtFunctionLibShared.DvdcTrim, StringComparison.OrdinalIgnoreCase)))
            {
                allForcePins.AddRange(measPins.SelectMany(x => x.ForceConditions).SelectMany(x => x.ForcePins).Where(forcePin => !allForcePins.Exists(s => s.PinName.Equals(forcePin.PinName))));
            }
            else
            {
                foreach (MeasPin measPin in measPins)
                {
                    string blockforcetype = GetForceTypeByMeasType(measPin.MeasType);
                    string meastypeComparsion = measPin.MeasType.Replace("Meas", "");
                    foreach (ForceCondition condition in measPin.ForceConditions)
                    {
                        foreach (ForcePin forcePin in condition.ForcePins)
                        {
                            //All the force pins those not force to the measure pins are "Premeas" ==>for VIR/VIF function
                            List<ForcePin> newforcePins = new List<ForcePin>
                                {
                                    forcePin.Copy()
                                };
                            List<ForcePin> allForcePinsBeforeDecomp = GetForcePinForPreMeasure(newforcePins, measPins, meastypeComparsion, blockforcetype, ref allForcePins);

                            // Check if force pin is group pin
                            if (allForcePinsBeforeDecomp != null)
                            {
                                if (allForcePinsBeforeDecomp.Count > 0)
                                {
                                    List<ForcePin> decompForcePins = DivideGroupPin(allForcePinsBeforeDecomp);
                                    List<ForcePin> allForcePinsAfterDecomp = GetForcePinForPreMeasure(decompForcePins, measPins, meastypeComparsion, blockforcetype, ref allForcePins);
                                    allForcePins.AddRange(allForcePinsAfterDecomp.Count != decompForcePins.Count ? allForcePinsAfterDecomp : allForcePinsBeforeDecomp);
                                }
                            }
                        }
                    }
                }
            }

            if (allForcePins.Count > 0)
            {
                foreach (ForcePin forcePin in allForcePins)
                {
                    if (forcePin.ForceLabelVoltages.Count > 0 && voltage != null &&
                                    !forcePin.ForceLabelVoltages.Contains(voltage))
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(forcePin.ForceInterPostMeas))
                    {
                        postMeasSeq.Add(forcePin.ForceInterPostMeas);
                    }

                }
            }
            postMeas.Add(postMeasSeq.All(string.IsNullOrEmpty) ? "" : string.Join(";", postMeasSeq));
        }

        public static string GetForceTypeByMeasType(string measpintype)
        {
            if (measpintype == MeasType.MeasV)
            {
                return "^I";
            }

            if (measpintype == MeasType.MeasI)
            {
                return "^V";
            }

            if (measpintype == MeasType.MeasR1 || measpintype == MeasType.MeasR2)
            {
                return "^(V|I)$";
            }

            if (measpintype == MeasType.MeasVdiff2) // force condition was Vp|Vn
            {
                return "^(V1P|V2P|V1N|V2N|V)$";
            }

            return "^Not_Define$";
        }

        public static List<ForcePin> DivideGroupPin(List<ForcePin> forcePinlst)
        {
            List<ForcePin> newforcePinlst = new List<ForcePin>();
            foreach (ForcePin forcePin in forcePinlst)
            {
                List<string> stringlist = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(forcePin.PinName);
                foreach (string pinname in stringlist)
                {
                    ForcePin newforcePin = forcePin.Copy();
                    newforcePin.PinName = pinname;
                    newforcePinlst.Add(newforcePin);
                }
            }
            return newforcePinlst;
        }

        private static List<ForcePin> GetForcePinForPreMeasure(List<ForcePin> forcePins, List<MeasPin> measPins, string measTypeComparsion, string blockForceType, ref List<ForcePin> allForcePinList)
        {
            List<ForcePin> allForcePins = allForcePinList;
            var newAllForcePins = new List<ForcePin>();
            //MeasE: Measure IO voltage with 2 different force current and calculate the difference to get calibrated voltage measurement (eg. one by one)
            foreach (ForcePin forcePin in forcePins)
            {
                bool isMeasTypeComparsion = forcePin.ForceType.Equals(measTypeComparsion, StringComparison.OrdinalIgnoreCase);
                bool isBlockForceType = Regex.IsMatch(forcePin.ForceType, blockForceType, RegexOptions.IgnoreCase);
                bool isMeasPinExists = measPins.Exists(g => g.PinName.Split(':').ToList().Exists(a => a.Equals(forcePin.PinName, StringComparison.OrdinalIgnoreCase)));
                if (!allForcePins.Exists(s => s.Equals(forcePin)))
                {
                    // The type of force and measure are the same => the force and measure can not the same (eg. measI & force is I)
                    if (isMeasTypeComparsion && !isMeasPinExists)
                    {
                        newAllForcePins.Add(forcePin);
                    }
                    // The type of force pin is not the opposite type of meas pin (eg. measI & force is not V or I)
                    else if (!isBlockForceType && !isMeasTypeComparsion)
                    {
                        newAllForcePins.Add(forcePin);
                    }
                    // The type of force pin is the opposite type of meas pin (eg. measI & force is V)
                    else if (isBlockForceType && !isMeasPinExists)
                    {
                        newAllForcePins.Add(forcePin);
                    }
                    else if (!forcePin.IsRestore)
                    {
                        newAllForcePins.Add(forcePin);
                    }
                }
            }

            return newAllForcePins;
        }

        public static string GetPreMeas(HardIpPattern pattern, HardIpInputData hardIpInputData, string testSequence = "", string voltage = "", string preMeasDelimiter = "|")
        {
            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            List<string> infoSequence = info.MeasSeqStr == "" ? GetMeasStrByPlan(pattern) : info.MeasSeqStr.Split(',').ToList();

            var preMeas = new List<string>();
            if (infoSequence.Count > 0)
            {
                var disconnectPinsFromIoPins = new List<string>();
                if (pattern.Pattern.GetLastPayload().ContainsIgnoreCase("HIZ") ||
                    pattern.Pattern.GetLastPayload().ContainsIgnoreCase("HIGHZ"))
                {
                    var iopins =
                        pattern.MeasPins.Where(
                            p => GetPinType(p.PinName).Equals("I/O", StringComparison.OrdinalIgnoreCase))
                            .Select(p => p.PinName)
                            .ToList();
                    disconnectPinsFromIoPins.AddRange(iopins);
                }

                for (int index = 1; index <= infoSequence.Count; index++)
                {

                    var measPins = pattern.MeasPins.Where(a => a.SequenceIndex == index).ToList();
                    var preMeasSeq = new List<string>();
                    var allForcePins = new List<ForcePin>();
                    //Collect all force pins in the same sequenceIndex
                    if (!(pattern.FunctionName.Equals(VbtFunctionLibShared.VifName, StringComparison.OrdinalIgnoreCase) ||
                          pattern.FunctionName.Equals(VbtFunctionLibShared.LcdMeas, StringComparison.OrdinalIgnoreCase) ||
                          pattern.FunctionName.Equals(VbtFunctionLibShared.LcdTrim, StringComparison.OrdinalIgnoreCase) ||
                          pattern.FunctionName.Equals(VbtFunctionLibShared.DvdcTrim3D, StringComparison.OrdinalIgnoreCase) ||
                    pattern.FunctionName.Equals(VbtFunctionLibShared.DvdcTrim, StringComparison.OrdinalIgnoreCase)))
                    {
                        allForcePins.AddRange(measPins.SelectMany(x => x.ForceConditions).SelectMany(x => x.ForcePins).Where(forcePin => !allForcePins.Exists(s => s.PinName.Equals(forcePin.PinName))));
                    }
                    else
                    {
                        if (pattern.OriMeasPins.Count == 1 && CheckByOriMeasPins(pattern, allForcePins))
                        {
                        }
                        else
                        {
                            CollectPreMeasForcePins(measPins, allForcePins, ref disconnectPinsFromIoPins, preMeasSeq);
                        }
                    }
                    allForcePins = allForcePins.GroupBy(p => p.PinName + p.ForceType).Select(p => p.First()).ToList();
                    if (allForcePins.Count > 0)
                    {
                        BuildPreMeasSeq(allForcePins, voltage, preMeasSeq);
                    }
                    if (preMeasSeq.All(string.IsNullOrEmpty))
                    {
                        preMeas.Add("");
                    }
                    else
                    {
                        preMeas.Add(string.Join(";", preMeasSeq));
                    }
                }
                if (preMeas.All(string.IsNullOrEmpty))
                {
                    return "";
                }

                string result = string.Join(preMeasDelimiter, preMeas);
                if (result.Length >= 6000)
                {
                    result = BuildPreMeasRegAssign(pattern, hardIpInputData, testSequence, preMeas);
                }
                return result;
            }
            return "";
        }

        private static void CollectPreMeasForcePins(List<MeasPin> measPins, List<ForcePin> allForcePins, ref List<string> disconnectPinsFromIoPins, List<string> preMeasSeq)
        {
            foreach (MeasPin measPin in measPins)
            {
                if (measPin.MeasType.Equals(MeasType.MeasVdiff2, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string blockforcetype = GetForceTypeByMeasType(measPin.MeasType);
                string meastypeComparsion = measPin.MeasType.Replace("Meas", "");

                var allForceConditionPins = measPin.ForceConditions.SelectMany(forcecondition =>
                    forcecondition.ForcePins.SelectMany(forcepin => TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(forcepin.PinName))).ToList();
                foreach (ForceCondition condition in measPin.ForceConditions)
                {
                    foreach (ForcePin forcePin in condition.ForcePins)
                    {
                        //All the force pins those not force to the measure pins are "Premeas" ==>for VIR/VIF function
                        List<ForcePin> newforcePins = new List<ForcePin>
                        {
                            forcePin.Copy()
                        };
                        List<ForcePin> allForcePinsBeforeDecomp = GetForcePinForPreMeasure(newforcePins, measPins, meastypeComparsion, blockforcetype, ref allForcePins);

                        // Check if force pin is group pin
                        if (allForcePinsBeforeDecomp != null)
                        {
                            if (allForcePinsBeforeDecomp.Count > 0)
                            {
                                List<ForcePin> decompForcePins = DivideGroupPin(allForcePinsBeforeDecomp);
                                List<ForcePin> allForcePinsAfterDecomp = GetForcePinForPreMeasure(decompForcePins, measPins, meastypeComparsion, blockforcetype, ref allForcePins);
                                bool isRestore = allForcePinsAfterDecomp.Any(x => !x.IsRestore); // NoRestore need the force conition syntax
                                allForcePins.AddRange(allForcePinsAfterDecomp.Count != decompForcePins.Count ? allForcePinsAfterDecomp :
                                    !isRestore && IsMeasPinInForcePin(forcePin.PinName, measPin.PinName, allForceConditionPins) && forcePin.Type != ForceConditionType.Others ? new List<ForcePin>()
                                        : allForcePinsBeforeDecomp);
                            }
                        }
                    }
                }
            }
            if (measPins.Exists(p => GetPinType(p.PinName).Equals("power", StringComparison.OrdinalIgnoreCase)))
            {
                disconnectPinsFromIoPins = disconnectPinsFromIoPins.Select(x => x.ToUpper()).Distinct().ToList();
                IEnumerable<ForcePin> disconnectPinsFromForceCondition = allForcePins.Where(x => x.ForceValue.Equals("DisconnectDigital", StringComparison.OrdinalIgnoreCase));
                var disconnectDecomposedPinsFromForceCondition = new List<string>();
                foreach (ForcePin item in disconnectPinsFromForceCondition)
                {
                    disconnectDecomposedPinsFromForceCondition.AddRange(TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(item.PinName).Select(x => x.ToUpper()));
                }

                if (disconnectDecomposedPinsFromForceCondition.Any() &&
                    !disconnectPinsFromIoPins.All(disconnectDecomposedPinsFromForceCondition.Contains))
                {
                    foreach (string digitalpin in disconnectPinsFromIoPins)
                    {
                        preMeasSeq.Add($"{digitalpin}:DisconnectDigital");
                    }
                }
            }
        }

        private static void BuildPreMeasSeq(List<ForcePin> allForcePins, string voltage, List<string> preMeasSeq)
        {
            var disableFrcPins = allForcePins.Where(p => p.ForceType.Equals("DISABLE_FRC", StringComparison.OrdinalIgnoreCase)).GroupBy(p => p.PinName).ToDictionary(p => p.Key, p => p.ToList());
            foreach (KeyValuePair<string, List<ForcePin>> disableFrcPin in disableFrcPins)
            {
                if (disableFrcPin.Value.Any(p => string.IsNullOrEmpty(p.ForceValue)))
                {
                    continue;
                }

                disableFrcPin.Value.ForEach(p => p.ForceType = "V");
                List<string> groupPins = TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(disableFrcPin.Key);
                groupPins.RemoveAll(p => Regex.IsMatch(p, "REFCLK", RegexOptions.IgnoreCase));
                allForcePins.Insert(0, new ForcePin { PinName = disableFrcPin.Key, ForceValue = "DISABLE_FRC", Type = ForceConditionType.Others });
                if (groupPins.Count == disableFrcPin.Value.Count)
                {
                    int interFrcPinIndex = 0;
                    foreach (ForcePin forceFrc in disableFrcPin.Value)
                    {
                        forceFrc.PinName = groupPins[interFrcPinIndex];
                        interFrcPinIndex++;
                    }
                }
            }
            foreach (ForcePin forcePin in allForcePins)
            {
                if (forcePin.ForceLabelVoltages.Count > 0 && voltage != null &&
                                !forcePin.ForceLabelVoltages.Contains(voltage))
                {
                    continue;
                }

                if (forcePin.Type == ForceConditionType.Normal)
                {

                    if (forcePin.ForceJob == "")
                    {
                        preMeasSeq.Add(forcePin.PinName + ":" + forcePin.ForceType + ":" + DataConvertor.ConvertForceValueToGlbSpec(forcePin, "", true) + (forcePin.IsRestore ? "" : ":NoRestore"));
                    }
                    else
                    {
                        preMeasSeq.Add(forcePin.PinName + ":" + forcePin.ForceType + ":" + DataConvertor.ConvertForceValueToGlbSpec(forcePin, "", true) + ":" + forcePin.ForceJob);
                    }
                }

                if (forcePin.Type == ForceConditionType.Others)
                {
                    if (forcePin.ForceJob == "")
                    {
                        preMeasSeq.Add(forcePin.PinName + ":" + forcePin.ForceValue + (forcePin.IsRestore ? "" : ":NoRestore"));
                    }
                    else
                    {
                        preMeasSeq.Add(forcePin.PinName + ":" + forcePin.ForceValue + ":" + forcePin.ForceJob);
                    }
                }
            }
        }

        private static string BuildPreMeasRegAssign(HardIpPattern pattern, HardIpInputData hardIpInputData, string testSequence, List<string> preMeas)
        {
            var hardIpRegAssign = new HardIpRegAssign();
            string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            hardIpRegAssign.SubBlockName = blockName + "_" + CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName);
            if (Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^dd_", RegexOptions.IgnoreCase))
            {
                hardIpRegAssign.SubBlockName += "_DD";
            }

            hardIpRegAssign.Type = RegisterAssignType.Interpose_PreMeas;
            if (!hardIpInputData.HardIpRegAssigns.Exists(x => x.SubBlockName.Equals(hardIpRegAssign.SubBlockName, StringComparison.CurrentCultureIgnoreCase) &&
                                                              x.Type == hardIpRegAssign.Type))
            {
                List<List<string>> regAssginList = new List<List<string>>();
                if (string.IsNullOrEmpty(testSequence))
                {
                    for (int index = 0; index < preMeas.Count; index++)
                    {
                        List<string> data = new List<string> { preMeas[index] };
                        regAssginList.Add(data);
                    }
                }
                else
                {
                    List<string> testSequenceArr = testSequence.Split(',').ToList();
                    if (testSequenceArr.Count == preMeas.Count)
                    {
                        for (int index = 0; index < preMeas.Count; index++)
                        {
                            List<string> data = new List<string>
                            {
                                testSequenceArr.ElementAt(index), preMeas[index]
                            };
                            regAssginList.Add(data);
                        }
                    }
                }
                hardIpRegAssign.RegAssignList = regAssginList;
                hardIpInputData.HardIpRegAssigns.Add(hardIpRegAssign);

            }
            return $"Reg_assign:{hardIpRegAssign.SubBlockName}";
        }

        private static bool CheckByOriMeasPins(HardIpPattern pattern, List<ForcePin> allForcePins)
        {
            List<ForcePin> newallForcePins = new List<ForcePin>();
            foreach (MeasPin measPin in pattern.OriMeasPins)
            {
                string blockforcetype = GetForceTypeByMeasType(measPin.MeasType);
                string meastypeComparsion = measPin.MeasType.Replace("Meas", "");
                foreach (ForceCondition condition in measPin.ForceConditions)
                {
                    foreach (ForcePin forcePin in condition.ForcePins)
                    {
                        //All the force pins those not force to the measure pins are "Premeas" ==>for VIR/VIF function
                        List<ForcePin> newforcePins = new List<ForcePin>
                        {
                            forcePin.Copy()
                        };
                        List<ForcePin> allForcePinsBeforeDecomp = GetForcePinForPreMeasure(newforcePins, pattern.OriMeasPins, meastypeComparsion, blockforcetype, ref allForcePins);
                        newallForcePins.AddRange(allForcePinsBeforeDecomp);

                    }
                }
            }
            return !newallForcePins.Any();
        }

        public static string GetMeasSequence(HardIpPattern pattern)
        {
            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            var measList = new Dictionary<int, int>();
            foreach (MeasPin pin in pattern.MeasPins)
            {

                int maxForceCnt = 1;
                if (pin.ForceConditions.Any())
                {
                    if (pin.MeasType == MeasType.MeasI || pin.MeasType == MeasType.MeasV)
                    {
                        maxForceCnt = pin.ForceConditions.SelectMany(x => x.ForcePins).Max(x => x.ForceCnt);
                    }
                    else if (pin.MeasType == MeasType.MeasR2)
                    {
                        maxForceCnt = pin.ForceConditions.SelectMany(x => x.ForcePins).Max(x => x.ForceCnt) / 2;
                    }

                    maxForceCnt = maxForceCnt < 1 ? 1 : maxForceCnt;
                }
                if (pin.MeasType != "MeasC" && pin.MeasType != MeasType.MeasLimit && pin.MeasType != MeasType.MeasCalc && !measList.ContainsKey(pin.SequenceIndex))
                {
                    measList.Add(pin.SequenceIndex, maxForceCnt);
                }
            }
            List<string> infoSequence = info.MeasSeqStr == "" ? GetMeasStrByPlan(pattern) : info.MeasSeqStr.Split(',').ToList();
            for (int i = 1; i <= infoSequence.Count; i++)
            {
                string mySeq = infoSequence[i - 1].Split('>')[0];
                if (!measList.TryGetValue(i, out int value))
                {
                    infoSequence[i - 1] = "N";
                }
                else
                {
                    mySeq = new StringBuilder().Insert(0, mySeq, value).ToString();
                    infoSequence[i - 1] = mySeq;
                }
            }
            return string.Join(",", infoSequence.ToArray()).ToUpper();
        }

        public static List<string> GetSeqlstFromMiscInfo(string miscInfo)
        {
            string testSeqSetting = miscInfo.Split(';').ToList().Find(s => Regex.IsMatch(s, HardIpConstData.RegTestSequence, RegexOptions.IgnoreCase));

            return testSeqSetting?.Split(':')[1].Split(',').ToList();
        }

        /// <summary>
        /// AutoGen only deal with the pattern start with dd,cz or pp
        /// if no pattern, use "No_patt" or specify instance name like "Instance:XXX" in pattern column
        /// </summary>
        /// <param name="patternName"></param>
        /// <returns></returns>
        public static bool IsValidPatName(string patternName)
        {
            return _regex.IsMatch(patternName) ||
                patternName.Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase) ||
                HardIpConstData.RegInsInPatt.IsMatch(patternName);
        }

        public static string GetIRangeByJob(string measureIRange)
        {

            string[] measList = Regex.Split(measureIRange, "[,+]");
            List<char> delimiter = measureIRange.ToCharArray().Where(s => s == ',' || s == '+').ToList();
            List<string> joblist = new List<string>();
            string[] jobValue;

            // Get array size and job list
            int meascnt = measList.Length;
            int jobcnt = 0;
            foreach (string pin in measList)
            {
                jobValue = pin.Split(';');
                jobcnt = jobValue.Length > jobcnt ? jobValue.Length : jobcnt;
                foreach (string job in jobValue)
                {
                    if (job.Split(':').Length > 1)
                    {
                        joblist.Add(job.Split(':')[0]);
                    }
                }
            }

            var modjoblist = joblist.Distinct().ToList();
            if (modjoblist.Any())
            {
                if (joblist.Count < meascnt)
                {
                    //modjoblist = new List<string> { "CP1", "CP2", "FT1", "FT2" };
                    //jobcnt = modjoblist.Count;
                }
                //Transfer data to array
                string[,] array = new string[meascnt, jobcnt];
                for (int i = 0; i < meascnt; i++)
                {
                    jobValue = measList[i].Split(';');
                    for (int j = 0; j < jobValue.Length; j++)
                    {
                        array[i, j] = jobValue[j].Split(':').Length > 1 ? jobValue[j].Split(':')[1] : jobValue[j].Split(':')[0];
                    }
                }

                //Convert to new IRange string
                string iRange = "";
                for (int j = 0; j < jobcnt; j++)
                {
                    iRange += modjoblist[j] + "=";
                    for (int i = 0; i < meascnt; i++)
                    {
                        if (measList[i].Split(';').Length > 1)
                        {
                            if (array[i, j] == null)
                            {
                                continue;
                            }

                            iRange += array[i, j] + (i >= delimiter.Count ? "" : delimiter[i].ToString());
                        }
                        else
                        {
                            iRange += array[i, 0] + (i >= delimiter.Count ? "" : delimiter[i].ToString());
                        }
                    }
                    iRange += ";";
                }
                iRange = iRange.Remove(iRange.Length - 1, 1);

                return iRange;
            }

            return measureIRange;
        }

        public static bool IsAllTheSameType(List<MeasPin> measPins, string measType)
        {
            return measPins.All(p => p.MeasType.Equals(measType, StringComparison.OrdinalIgnoreCase));
        }

        public static string CheckInfoByStoreName(string info, string storeName, char sign, bool isTestSeq = false)
        {
            if (string.IsNullOrEmpty(storeName))
            {
                return info;
            }

            try
            {
                var result = new List<string>();
                List<string> storeNameList = storeName.Split('+').ToList();
                if (isTestSeq)
                {
                    int i = 0;
                    if (info.Split(',').Length == 1)
                    {
                        string seqTmp = info;
                        for (int k = 0; k < storeNameList[i].Split(':').Length - 1; k++)
                        {
                            seqTmp += info;
                        }
                        result.Add(seqTmp);
                    }
                    else
                    {
                        foreach (string seq in info.Split(','))
                        {
                            string seqTmp = seq;
                            for (int k = 0; k < storeNameList[i].Split(':').Length - 1; k++)
                            {
                                seqTmp += seq;
                            }
                            result.Add(seqTmp);
                            i++;
                        }
                    }
                    return string.Join(",", result);
                }
                else
                {
                    int i = 0;
                    if (info.Split(sign).Length == 1)
                    {
                        if (storeNameList[i].Contains(":"))
                        {
                            result.Add(info.Replace(",", ":"));
                        }
                        else
                        {
                            result.Add(info);
                        }
                    }
                    else
                    {
                        foreach (string seqinfo in info.Split(sign))
                        {
                            if (storeNameList[i].Contains(":"))
                            {
                                result.Add(seqinfo.Replace(",", ":"));
                            }
                            else
                            {
                                result.Add(seqinfo);
                            }

                            i++;
                        }
                    }
                    return string.Join(sign.ToString(), result);
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return info;
        }

        public static string GetInstrumentInfo(HardIpPattern pattern, string item)
        {
            var itemSeq = new List<string>();
            string regItem = $@"{item}\s*=\s*(?<value>[\w,#]+)";
            IEnumerable<MeasPin> pinsPrefilter =
                pattern.MeasPins.Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase) &&
                                !p.MeasType.Equals(MeasType.MeasLimit, StringComparison.OrdinalIgnoreCase));
            if (pinsPrefilter.Count() == 0)
            {
                return "";
            }

            for (int seqindex = 1; seqindex <= pinsPrefilter.Max(p => p.SequenceIndex); seqindex++)
            {
                IEnumerable<MeasPin> pins = pattern.MeasPins.Where(p => seqindex == p.SequenceIndex);
                var setups = pins.GroupBy(p => p.RfInstrumentSetup).ToDictionary(p => p.Key, p => p.ToList());
                var items = new List<string>();
                foreach (string setup in setups.Keys)
                {
                    string selectItem = "";
                    foreach (string setupInfo in setup.Split('$'))
                    {
                        if (Regex.IsMatch(setupInfo, regItem, RegexOptions.IgnoreCase))
                        {
                            selectItem = Regex.Match(setupInfo, regItem, RegexOptions.IgnoreCase).Groups["value"].ToString();
                            break;
                        }
                    }
                    items.Add(selectItem);
                }
                itemSeq.Add(string.Join("+", items));
            }
            return string.Join("|", itemSeq);

        }

        public static ForceConditionType GetForceMode(List<MeasPin> measPins)
        {
            var forcePins = measPins.SelectMany(x => x.ForceConditions.SelectMany(y => y.ForcePins)).ToList();
            if (forcePins.Exists(x => x.ForceType.Equals("V", StringComparison.CurrentCultureIgnoreCase)))
            {
                return ForceConditionType.FvMode;
            }

            if (forcePins.Exists(x => x.ForceType.Equals("I", StringComparison.CurrentCultureIgnoreCase)))
            {
                return ForceConditionType.FiMode;
            }

            return ForceConditionType.Normal;
        }

        private static List<string> GetForcePairValueDotNet(List<string> forceList)
        {
            if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                return forceList;
            }

            bool waitMerge = false;
            string firstPinValue1 = "";
            string firstPinValue2 = "";
            var newForceList = new List<string>();
            foreach (string force in forceList)
            {
                string[] valuePair = force.Split('&');
                if (valuePair.Length == 2)
                {
                    if (waitMerge)
                    {
                        string secondPinValue1 = valuePair[0];
                        string secondPinValue2 = valuePair[1];
                        newForceList.Add(string.Format($"{firstPinValue1},{secondPinValue1}&{firstPinValue2},{secondPinValue2}"));
                        waitMerge = false;
                        firstPinValue1 = "";
                        firstPinValue2 = "";
                    }
                    else
                    {
                        firstPinValue1 = valuePair[0];
                        firstPinValue2 = valuePair[1];
                        waitMerge = true;
                    }
                }
                else
                {
                    newForceList.Add(force);
                }
            }
            return newForceList;
        }

        public static string GetStoreNameOri(HardIpPattern pattern)
        {
            var storeNameList = new List<string>();
            var pinsInSeq = pattern.MeasPins.GroupBy(p => p.SequenceIndex).ToDictionary(p => p.Key, p => p.ToList());

            foreach (KeyValuePair<int, List<MeasPin>> pins in pinsInSeq)
            {

                if (pins.Key < 1)
                {
                    continue;
                }

                //concern CP:/FT: with common case
                var storeNameSeqList = pins.Value.Where(p => !p.PinName.Contains("=")).Select(p => p.CusStr).ToList();
                storeNameSeqList.AddRange(pins.Value.Where(p => p.PinName.Contains("CP=")).Select(p => p.CusStr));
                if (storeNameSeqList.Distinct().Count(p => !string.IsNullOrEmpty(p)) == storeNameSeqList.Count)
                {
                    storeNameList.Add(string.Join(":", storeNameSeqList));
                }
                else
                {
                    storeNameList.Add(storeNameSeqList.FirstOrDefault(p => !string.IsNullOrEmpty(p)));
                }
            }

            string storeName = storeNameList.Any(p => !string.IsNullOrEmpty(p)) ? string.Join("+", storeNameList) : "";
            return storeName;
        }
    }
}
