using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;

using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetVifValue : SetValueBase
    {
        public SetVifValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            string cPin = SearchInfo.GetMeasCPins(pattern, HardIpService.GetHardIpInfo(pattern));
            //patset
            function.SetParamValue("patset", pattern.Pattern.IsMultiple() ?
                string.Join(",", pattern.Pattern.InstancePayloadName) : pattern.Pattern.GetInstancePatternName());
            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            //Cpu_flag_A
            HardIpInfo infoCpuFlag = HardIpService.GetHardIpInfo(pattern.Pattern.GetLastPayload());
            function.SetParamValue("CPUA_Flag_In_Pat", SearchInfo.GetCpuflag(infoCpuFlag, pattern));

            #region Defalut value
            //TestLimitPerPin
            //function.SetParamValue("TestLimitPerPin_VFI", string.Join("", SearchInfo.GetTestLimitPerMeasType(pattern).Values.ToList()));
            //DisableComparePins
            //function.SetParamValue("DisableComparePins", "-1");
            //MeasF_Interval
            if (!string.IsNullOrEmpty(pattern.BlockType))
            {
                function.SetParamValue("MeasF_Interval", "0.01");
            }

            //MeasF_EventSourceWithTerminationMode
            if (!string.IsNullOrEmpty(pattern.BlockType))
            {
                function.SetParamValue("MeasF_EventSourceWithTerminationMode", "2");
            }
            else
            {
                function.SetParamValue("MeasF_EventSourceWithTerminationMode", "0");
            }

            //MeasF_ThresholdPercentage
            function.SetParamValue("MeasF_ThresholdPercentage", "0.5");
            //MeasF_WaitTime
            //function.SetParamValue("MeasF_WaitTime", "0.01");
            #endregion

            #region RegisterAssignment
            //DigSrc_Assignment: Use "Register Assignment" value in test plan
            //function.ArgList[25] = pattern.RegisterAssignment;
            bool isAddSweep = false;
            function.SetParamValue("DigSrc_Assignment", pattern.ProcessSweepData(ref isAddSweep));
            SetDigSrcEquationAndSize(pattern, function, info);

            //DigSrc_Pin
            function.SetParamValue("DigSrc_Pin", SearchInfo.GetSrcPin(info));
            #endregion

            #region MeasC
            //DigCap_Pin: MeasC pin in Test Plan, Like "MeasC Pin = Pout" ===> Pout
            //function.ArgList[17] = cPin;
            function.SetParamValue("DigCap_Pin", cPin);
            //DigCap_DataWidth:  Get from "Cap Bit Str" in patInfo file. Like "wdr14_10+wdr23_10" ===> 10
            //function.ArgList[18] = Regex.Match(info.CapBitStr, @"^wdr\d+_(?<num>(\d+)).*").Groups["num"].ToString();
            function.SetParamValue("DigCap_DataWidth", SearchInfo.GetDigDataWidth(info.CapBitStr));
            //DigCap_Sample_Size: Get from "Cap Bit" in patInfo file
            //function.ArgList[19] = info.CapBit.ToString();
            function.SetParamValue("DigCap_Sample_Size", info.CapBit.ToString("G15"));

            //CUS_Str_DigCapData
            HardIpInfo hardIpInfo = HardIpService.GetHardIpInfo(pattern);
            function.SetParamValue("CUS_Str_DigCapData", SearchInfo.GetCusStrDigCapData(pattern, hardIpInfo));
            SetDigCapPostProcessing(pattern, function);

            //CUS_Str_MainProgram
            SetMainProgramCustomString(pattern, function);
            if (isAddSweep && pattern.ConditionIndex == 0)
            {
                pattern.DspFunction.RemoveAt(pattern.DspFunction.Count - 1);
            }

            string measSeq = GetMeasSeq(SearchInfo.GetMeasSequence(pattern).ToUpper());
            string storeNameOri = "";
            string measStoreName = DataConvertor.SortCpFtPin(SearchInfo.GetStoreName(pattern, HardIpInputData, ref storeNameOri, measSeq));
            function.SetParamValue("Meas_StoreName", measStoreName);
            #endregion

            if (measSeq.Contains("R") || measSeq.Contains("Z"))
            {
                function.SetParamValue("RAK_Flag", "1");
            }

            string forceV = SearchInfo.GetForceV(pattern, info, voltage);
            string forceI = SearchInfo.GetForceI(pattern, info, voltage);
            function.SetParamValue("ForceV_Val", CheckAddSymbol(SearchInfo.CheckInfoByStoreName(forceV, storeNameOri, '|')));
            function.SetParamValue("ForceI_Val", CheckAddSymbol(SearchInfo.CheckInfoByStoreName(forceI, storeNameOri, '|')));
            //Measure Sequence 
            string testSequence = SearchInfo.CheckInfoByStoreName(measSeq, storeNameOri, ',', true);
            function.SetParamValue("TestSequence", testSequence);
            //_WaitTime
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                function.SetParamValue("MeasF_WaitTime", GetWaitTime(testSequence, "F"));
                function.SetParamValue("MeasI_WaitTime", GetWaitTime(testSequence, "I"));
                function.SetParamValue("UVI80_MeasV_WaitTime", GetWaitTime(testSequence, "V"));
            }
            #region Meas Pins
            int seqCount = info.NewInfo != null ? info.NewInfo.SeqInfo.Count :
                info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
            if (seqCount > 0)
            {
                SetMeasPinsArgs(pattern, function, voltage, info, seqCount, storeNameOri, testSequence);
            }
            #endregion


            //if MeasI2,set SpecialCalcValSetting=3
            if (pattern.SpecialMeasType.Equals(MeasType.MeasVdiff))
            {
                function.SetParamValue("SpecialCalcValSetting", "4");
            }
            else if (pattern.SpecialMeasType.Equals(MeasType.MeasVocm))
            {
                function.SetParamValue("SpecialCalcValSetting", "9");
            }

            //Interpose_PrePat,Premeas,PostTest
            function.SetParamValue("Interpose_PrePat", SearchInfo.GetPrePat(pattern, voltage));
            function.SetParamValue("Interpose_PreMeas", SearchInfo.GetPreMeas(pattern, HardIpInputData, testSequence, voltage));
            function.SetParamValue("Interpose_PostMeas", SearchInfo.GetPostMeas(pattern));
            function.SetParamValue("Interpose_PostTest", pattern.InterposePostTest);

            //Calc_Eqn and storeName
            function.SetParamValue("Calc_Eqn", DataConvertor.ConvertValueSpec(pattern.CalcEqn));

            if (pattern.MiscInfo.Split(';').Any(
                p => p.Split(':')[0].Equals(HardIpConstData.InstSpecialSetup, StringComparison.OrdinalIgnoreCase)))
            {
                string specialsetup = pattern.MiscInfo.Split(';').FirstOrDefault(
                    p => p.Split(':')[0].Equals(HardIpConstData.InstSpecialSetup, StringComparison.OrdinalIgnoreCase));
                function.SetParamValue(HardIpConstData.InstSpecialSetup, GetInstSpecialSetup(specialsetup));
            }

            if (SearchInfo.GetFlagSingleLimit(pattern, voltage))
            {
                function.SetParamValue("Flag_SingleLimit", "TRUE");
            }

            //SweepCode
            function.SetParamValue("DigSrc_FlowForLoopIntegerName", HardIpService.GetFlowForLoopIntegerName(pattern, HardIpInputData));
            function.SetParamValue("MSB_First_Flag", info.IsMsbFirst());
        }

        private void SetMeasPinsArgs(HardIpPattern pattern, Function function, string voltage, HardIpInfo info, int seqCount,
            string storeNameOri, string testSequence)
        {
            #region Gets measPins from TestSequence

            var measVPins = new List<string>();
            var measIPins = new List<string>();
            var measFPins = new List<string>();
            //var measDPins = new List<string>();
            var measFdiffPins = new List<string>();

            for (int sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
            {
                CollectMeasPinsForSequence(pattern, sequenceIndex, measVPins, measIPins, measFPins, measFdiffPins);
            }
            #endregion

            //MeasureV_PinS
            if (measVPins.Any(p => !string.IsNullOrEmpty(p)))
            {
                string measV = FormatMeasPinValue(function, measVPins, storeNameOri);
                if (measV.Length >= 6000)
                {
                    measV = WriteMeasPinToRegAssign(pattern, testSequence, measVPins, "MeasV");
                }
                function.SetParamValue("MeasV_PinS", measV);
            }
            //MeasureF_PinS_SingleEnd
            if (measFPins.Any(p => !string.IsNullOrEmpty(p)))
            {
                string measF = FormatMeasPinValue(function, measFPins, storeNameOri);
                if (measF.Length >= 6000)
                {
                    measF = WriteMeasPinToRegAssign(pattern, testSequence, measFPins, "MeasF");
                }
                function.SetParamValue("MeasF_PinS_SingleEnd", measF);
            }
            //MeasureF_PinS_Differtial
            if (measFdiffPins.Any(p => !string.IsNullOrEmpty(p)))
            {
                string measFDiff = FormatMeasPinValue(function, measFdiffPins, storeNameOri);
                function.SetParamValue("MeasF_PinS_Differential", measFDiff);
            }
            //MeasureI_pinS
            if (measIPins.Any(p => !string.IsNullOrEmpty(p)))
            {
                //function.SetParamValue("MeasI_pinS", DataConvertor.SortCpFtPin(SearchInfo.CheckInfoByStoreName(DataConvertor.RemoveDummyPlusSign(string.Join("+", measIPins)), storeNameOri, '+')));
                string measI = FormatMeasPinValue(function, measIPins, storeNameOri);
                if (measI.Length >= 6000)
                {
                    measI = WriteMeasPinToRegAssign(pattern, testSequence, measIPins, "MeasI");
                }
                function.SetParamValue("MeasI_pinS", measI);

                //MeasI_Range
                //currentRange = SearchInfo.GetIRangeByJob(currentRange);
                //function.SetParamValue("MeasI_Range", DataConvertor.RedefineRange(pattern, info, SearchInfo.CheckInfoByStoreName(SearchInfo.GetIRange(pattern, info, voltage), storeNameOri, '|')));
                function.SetParamValue("MeasI_Range", SearchInfo.CheckInfoByStoreName(SearchInfo.GetIRange(pattern, info, voltage), storeNameOri, '|'));
            }
        }

        private void SetDigSrcEquationAndSize(HardIpPattern pattern, Function function, HardIpInfo info)
        {
            //DigSrc_Equation: From patInfo file "Send Bit Name"
            var equList = new List<string>();
            if (pattern.Pattern.IsMultiple())
            {
                foreach (string patternName in pattern.Pattern.PatternSetList.SelectMany(p => p).ToList())
                {
                    HardIpPattern pat = pattern.BurstPatterns.FirstOrDefault(x => x.Pattern.RealPatternName.Equals(patternName));
                    if (pat == null)
                    {
                        equList.Add("");
                    }
                    else
                    {
                        equList.Add(pat.DigSrcEquation);
                    }
                }
                pattern.DigSrcEquation = string.Join("|", equList);
            }

            function.SetParamValue("DigSrc_Equation", pattern.DigSrcEquation);
            //DigSrc_Sample_Size: Get from "Send Bit" in patInfo file, Like Send Bit: 160  ===> 160
            //function.ArgList[23] = info.SendBit.ToString();
            if (pattern.Pattern.IsMultiple())
            {
                var sendbit = new List<string>();
                foreach (string pat in pattern.Pattern.PatternSetList.SelectMany(p => p).ToList())
                {
                    HardIpInfo multiInfo = HardIpService.GetHardIpInfo(pat);
                    sendbit.Add(multiInfo.SendBit);
                }
                if (sendbit.Any(p => !string.IsNullOrEmpty(p)))
                {
                    info.SendBit = string.Join("|", sendbit);
                }
            }
            function.SetParamValue("DigSrc_Sample_Size", info.SendBit);
            //DigSrc_DataWidth: Get from "Send Bit Str" in patInfo file. Like wdr0_16+wdr1_16+wdr2_16 ===> 16
            //function.ArgList[22] =
            if (pattern.Pattern.IsMultiple())
            {
                var datawidth = new List<string>();
                foreach (string pat in pattern.Pattern.PatternSetList.SelectMany(p => p).ToList())
                {
                    HardIpInfo multiInfo = HardIpService.GetHardIpInfo(pat);
                    datawidth.Add(SearchInfo.GetDigDataWidth(multiInfo.SendBitStr, "0"));
                }
                if (datawidth.Any(p => !string.IsNullOrEmpty(p)))
                {
                    function.SetParamValue("DigSrc_DataWidth", string.Join("|", datawidth));
                }
            }
            else
            {
                function.SetParamValue("DigSrc_DataWidth", SearchInfo.GetDigDataWidth(info.SendBitStr, "0"));
            }
        }

        private void SetDigCapPostProcessing(HardIpPattern pattern, Function function)
        {
            string digCapDataPostProcessing = pattern.MiscInfo.Split(';').FirstOrDefault(x => x.StartsWith("DigCapData_PostProcessing", StringComparison.OrdinalIgnoreCase));
            string postProcess = pattern.MiscInfo.Split(';').FirstOrDefault(x => x.StartsWith("Post_Process", StringComparison.OrdinalIgnoreCase));

            if (LocalSpecs.Options.Device != EnumDevice.RF && !string.IsNullOrEmpty(digCapDataPostProcessing))
            {
                function.SetParamValue("CUS_Str_DigCapData", digCapDataPostProcessing.Replace(":", ","));
            }
            else if (LocalSpecs.Options.Device == EnumDevice.RF && !string.IsNullOrEmpty(postProcess))
            {
                function.SetParamValue("CUS_Str_DigCapData", postProcess.Replace(":", ","));
            }
        }

        private void SetMainProgramCustomString(HardIpPattern pattern, Function function)
        {
            if (pattern.Pattern.IsMultiple())
            {
                var dspList = new List<string>();
                foreach (string patternName in pattern.Pattern.PatternSetList.SelectMany(p => p).ToList())
                {
                    HardIpPattern pat = pattern.BurstPatterns.FirstOrDefault(x => x.Pattern.RealPatternName.Equals(patternName));
                    if (pat == null)
                    {
                        dspList.Add("");
                    }
                    else
                    {
                        dspList.Add(string.Join(";", pat.DspFunction));
                    }
                }
                if (dspList.Distinct().Count() == 1)
                {
                    function.SetParamValue("CUS_Str_MainProgram", dspList.First());
                }
                else if (dspList.Distinct().Count() > 1)
                {
                    function.SetParamValue("CUS_Str_MainProgram", string.Join("|", dspList));
                }
            }
            else
            {
                function.SetParamValue("CUS_Str_MainProgram", string.Join(";", pattern.DspFunction.Distinct()));
            }
        }

        private string FormatMeasPinValue(Function function, List<string> measPins, string storeNameOri)
        {
            string value = DataConvertor.SortCpFtPin(
                SearchInfo.CheckInfoByStoreName(string.Join("+", measPins), storeNameOri, '+'));
            if (function.FunctionName.ToLower() == VbtFunctionLibShared.VifName.ToLower() ||
                function.FunctionName.ToLower() == VbtFunctionLibShared.HardIpmtdTest.ToLower())
            {
                return DataConvertor.RemoveDummyPlusSign(value);
            }

            return value.Trim(',');
        }

        private void CollectMeasPinsForSequence(HardIpPattern pattern, int sequenceIndex, List<string> measVPins, List<string> measIPins, List<string> measFPins, List<string> measFdiffPins)
        {
            var seqPins = pattern.MeasPins.Where(p => p.SequenceIndex == sequenceIndex).ToList();
            bool isMeasR = seqPins.Exists(x => x.MeasType == MeasType.MeasR1 || x.MeasType == MeasType.MeasR2);
            ForceConditionType mode = isMeasR ? SearchInfo.GetForceMode(seqPins) : ForceConditionType.Normal;
            string selVType =
                seqPins.Exists(p => p.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase)) ? MeasType.MeasVdiff :
                MeasType.MeasV;
            bool isPinGroup = seqPins.Exists(p => TestProgram.IgxlWorkBk.PinMapPair.Value.IsGroupExist(p.PinName));
            var measIInter = new List<MeasPin>();
            var measVInter = new List<MeasPin>();

            if (mode == ForceConditionType.FiMode) //MeasR Force I mode
            {
                measIInter =
                seqPins.Where(
                    p => p.MeasType.Equals(MeasType.MeasI, StringComparison.OrdinalIgnoreCase)).ToList();
                measVInter = seqPins.Where(p =>
                p.MeasType.Equals(selVType, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase) || p.MeasType.Equals(MeasType.MeasR1, StringComparison.OrdinalIgnoreCase) ||
                        p.MeasType.Equals(MeasType.MeasR2, StringComparison.OrdinalIgnoreCase)
                ).ToList();
            }
            else if (mode == ForceConditionType.FvMode) // MeasR Force V mode
            {
                measIInter =
                seqPins.Where(
                    p => p.MeasType.Equals(MeasType.MeasI, StringComparison.OrdinalIgnoreCase)
                         || p.MeasType.Equals(MeasType.MeasR1, StringComparison.OrdinalIgnoreCase) ||
                         p.MeasType.Equals(MeasType.MeasR2, StringComparison.OrdinalIgnoreCase)).ToList();
                measVInter = seqPins.Where(p =>
                p.MeasType.Equals(selVType, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase)
                ).ToList();
            }
            else
            {
                measIInter =
                seqPins.Where(p =>
                p.MeasType.Equals(MeasType.MeasI, StringComparison.OrdinalIgnoreCase) ||
                p.MeasType.Equals(MeasType.MeasVdiff2, StringComparison.OrdinalIgnoreCase)
                ).ToList();

                measVInter = seqPins.Where(
                    p => p.MeasType.Equals(selVType, StringComparison.OrdinalIgnoreCase) ||
                    p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase)
                    ).ToList();
            }
            var measFInter = seqPins.Where(
                p => p.MeasType.Equals(MeasType.MeasF, StringComparison.OrdinalIgnoreCase) || p.MeasType.Equals(MeasType.MeasD)).ToList();
            var measFdiffInter = seqPins.Where(
                p => p.MeasType.Equals(MeasType.MeasFdiff, StringComparison.OrdinalIgnoreCase)).ToList();

            measVPins.Add(isPinGroup ? string.Join(",", RearrangePin(measVInter)) :
                string.Join(",", measVInter.Select(p =>
                p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase) ?
                p.PinName : SearchInfo.GenDiffGroupName(p.PinName, false))));

            if (measIInter.Any(p => p.MeasType == MeasType.MeasVdiff2) && measIInter.SelectMany(x => x.ForceConditions).Any()) //Vdiff2 meas pins from force for single pin limit by Orange/HungYu
            {
                measIPins.Add(string.Join(",", measIInter.SelectMany(p => p.ForceConditions.SelectMany(q => q.ForcePins.Select(s => s.PinName))).Distinct()));
            }
            else
            {
                measIPins.Add(string.Join(",", measIInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, false))));
            }

            measFPins.Add(string.Join(",", measFInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, true))));
            measFdiffPins.Add(string.Join(",", measFdiffInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, true))));
        }

        internal virtual string GetMeasSeq(string measSeq)
        {
            var list = new List<string>();
            foreach (string seq in measSeq.Split(','))
            {
                if (seq.Equals("VDIFF2"))
                {
                    list.Add(seq);
                }
                else
                {
                    list.Add(seq.Replace("FDIFF", "F").Replace("R1", "R").Replace("IDIFF", "I").Replace("R2", "Z").Replace("VDIFF", "V"));
                }
            }

            return string.Join(",", list);

        }

        internal List<string> RearrangePin(List<MeasPin> pins)
        {
            var forcePins = pins.SelectMany(p => p.ForceConditions).SelectMany(p => p.ForcePins).Select(p => p.PinName).Distinct().ToList();
            var measPins = pins.Select(p => p.PinName).Distinct().ToList();
            if (forcePins.Count == 0)
            {
                return measPins;
            }

            if (!forcePins.Except(measPins).Any() && !measPins.Except(forcePins).Any())
            {
                return measPins;
            }

            var totalMeasPins = new List<string>();
            foreach (string meas in measPins)
            {
                totalMeasPins.AddRange(TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(meas));
            }
            var totalForcePins = new List<string>();
            foreach (string force in forcePins)
            {
                totalForcePins.AddRange(TestProgram.IgxlWorkBk.PinMapPair.Value.DecompGroups(force));
            }
            if (!totalForcePins.Except(totalMeasPins).Any() &&
                !totalMeasPins.Except(totalForcePins).Any())
            {
                return forcePins;
            }
            return pins.Select(p => p.PinName).ToList();
        }

        internal string WriteMeasPinToRegAssign(HardIpPattern pattern, string testSequence, List<string> pinList, string type)
        {
            var hardIpRegAssign = new HardIpRegAssign();
            string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
            hardIpRegAssign.SubBlockName = blockName + "_" + CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName);
            if (Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^dd_", RegexOptions.IgnoreCase))
            {
                hardIpRegAssign.SubBlockName += "_DD";
            }

            switch (type)
            {
                case "MeasV":
                    hardIpRegAssign.Type = RegisterAssignType.MeasV_PinS;
                    break;
                case "MeasI":
                    hardIpRegAssign.Type = RegisterAssignType.MeasI_PinS;
                    break;
                case "MeasF":
                    hardIpRegAssign.Type = RegisterAssignType.MeasF_PinS_SingleEnd;
                    break;
            }

            if (!HardIpInputData.HardIpRegAssigns.Exists(x => x.SubBlockName.Equals(hardIpRegAssign.SubBlockName, StringComparison.CurrentCultureIgnoreCase) &&
                                                              x.Type == hardIpRegAssign.Type))
            {
                List<List<string>> regAssginList = new List<List<string>>();
                List<string> testSequenceArr = testSequence.Split(',').ToList();
                if (testSequenceArr.Count == pinList.Count)
                {
                    for (int index = 0; index < pinList.Count; index++)
                    {
                        string measPin = pinList[index];
                        List<string> data = new List<string> { testSequenceArr.ElementAt(index) };
                        data.AddRange(new List<string> { measPin });
                        regAssginList.Add(data);
                    }
                }

                hardIpRegAssign.RegAssignList = regAssginList;
                HardIpInputData.HardIpRegAssigns.Add(hardIpRegAssign);

            }
            return $"Reg_assign:{hardIpRegAssign.SubBlockName}";
        }

        internal string CheckAddSymbol(string input)
        {
            if (Regex.IsMatch(input, "[:+]", RegexOptions.IgnoreCase))
            {
                List<string> sgmts = Regex.Split(input, "[:+]", RegexOptions.IgnoreCase).ToList();
                if (sgmts.All(p => double.TryParse(p, out double _)))
                {
                    return "@" + input;
                }
            }
            return input;
        }

        internal string GetWaitTime(string allseq, string speseq)
        {
            if (!allseq.Contains(speseq))
            {
                return "";
            }

            List<string> resseq = new List<string>();
            foreach (string seq in allseq.Split(','))
            {
                if (seq == speseq)
                {
                    resseq.Add("0.01");
                }
                else
                {
                    resseq.Add("");
                }
            }
            return string.Join("+", resseq);
        }

        internal string GetInstSpecialSetup(string setup)
        {
            string result = "";
            if (setup.Contains(":"))
            {
                string value = setup.Split(':')[1].Trim();
                if (int.TryParse(value, out int _))
                {
                    result = value;
                }
                else
                {
                    if (HardIpInputData.ConfigData.InstSpecialSetting.TryGetValue(value, out string value1))
                    {
                        result = value1;
                    }
                    else
                    {
                        result = "KeyNotFoundInSetting";
                    }
                }
            }
            return result;
        }
    }
}
