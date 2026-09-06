using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting;
using Automation.GenerateIgxl.Wireless.DVDC.WirelessConst;
using Automation.InputManager.Data;
using Automation.Utility.HardIP;

using LogLib.Utility;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Wireless.DVDC.InstanceParameterSetting
{
    [ExcludeFromCodeCoverage]
    public class SetRfTrimValue : SetValueBase
    {
        private string _doFinalVerification = "FALSE";
        private string _doTrimming = "TRUE";
        private readonly List<string> _calclist = new List<string>();
        private readonly List<string> _calcStorelist = new List<string>();
        private readonly List<string> _measStoreName = new List<string>();
        private readonly List<string> _types = new List<string>();
        private string _sweepRange = "";

        public SetRfTrimValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {

            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            string payload = pattern.Pattern.GetLastPayload();

            function.SetParamValue("patset", payload);
            //function.SetParamValue("DsscSetup", pattern.WirelessData.RegisterAssignment);
            function.SetParamValue("DsscSetup", payload);
            string rfSetupName = SearchInfo.GetInstrumentInfo(pattern, "RFInstSetup");
            function.SetParamValue("RFInstSetup", rfSetupName);
            if (Regex.IsMatch(rfSetupName, "UWS|UWM|LXS|LXM", RegexOptions.IgnoreCase))
            {
                function.SetParamValue("InterposePostRst", "setUWDisconnectPinFlag(TRUE)");
            }

            SetMeasurement(pattern);
            /*                

                function.SetParamValue(WirelessConstData.DigSrcpin, SearchInfo.GetSrcPin(info));
                function.SetParamValue(WirelessConstData.DigSrcEquation, info.SendBitName);
                function.SetParamValue(WirelessConstData.DigSrcAssignment,
                    AppendRegisterBitWidth(info, pattern.WirelessData.RegisterAssignment));
             */

            GetTrimType(pattern.WirelessData.TrimType);
            function.SetParamValue(WirelessConstData.TrimRegName, pattern.WirelessData.TrimFuseName);
            if (!Regex.IsMatch(pattern.WirelessData.TrimTarget, "^=") &&
                Regex.IsMatch(pattern.WirelessData.TrimTarget, @".*[\*\+\^\-].*"))
            {
                pattern.WirelessData.TrimTarget = "=" + pattern.WirelessData.TrimTarget;
            }

            function.SetParamValue(WirelessConstData.TrimTarget, pattern.WirelessData.TrimTarget.Replace("&", "::"));
            function.SetParamValue(WirelessConstData.DigSrcpin, SearchInfo.GetSrcPin(info));
            function.SetParamValue(WirelessConstData.DigSrcEquation, info.SendBitName);
            function.SetParamValue(WirelessConstData.DigSrcAssignment,
                AppendRegisterBitWidth(info, pattern.WirelessData.RegisterAssignment));
            //DigCapPin	DigCapBits	DigCapData_PostProcessing
            function.SetParamValue("DigCapPin", info.CapPinName);
            //DigCap_Sample_Size: Get from "Cap Bit" in patInfo file
            function.SetParamValue("DigCapBits", info.CapBit == 0 ? "" : info.CapBit.ToString("G15"));

            //CUS_Str_DigCapData
            HardIpInfo hardIpInfo = HardIpService.GetHardIpInfo(pattern);
            function.SetParamValue("DigCapData_PostProcessing", SearchInfo.GetCusStrDigCapData(pattern, hardIpInfo));
            function.SetParamValue("SeqCnt", _types.Count.ToString());
            function.SetParamValue("TestType", string.Join("|", _types));
            function.SetParamValue("CalcEquName", string.Join("|", _calclist));
            function.SetParamValue("CalcStoreName", string.Join("|", _calcStorelist));
            function.SetParamValue("Meas_StoreName", string.Join("|", _measStoreName));
            function.SetParamValue(WirelessConstData.SweepRange, _sweepRange);
            function.SetParamValue("PrintDebugInfo", "TRUE");
            if (pattern.WirelessData.IsNeedPostBurn)
            {
                _doTrimming = "FALSE";
                _doFinalVerification = "TRUE";
            }
            else
            {
                _doTrimming = "TRUE";
                _doFinalVerification = "TRUE";
            }
            function.SetParamValue(WirelessConstData.DoTrimming, _doTrimming);
            function.SetParamValue(WirelessConstData.DoFinalVerification, _doFinalVerification);
            function.SetParamValue("DoApplyLevelsTiming", "True");
            SetInterPoseFunc(pattern, function);
            MergeForceItemsToInterPosePreMeas(function, pattern);
        }

        private void MergeForceItemsToInterPosePreMeas(Function function, HardIpPattern pattern)
        {
            try
            {
                string[] curInterPoseInfo = function.GetParamValue("InterPosePreMeas").Split('|');
                List<string> preMeasStr = ProcessSocPreMeasToRf(SearchInfo.GetPreMeas(pattern, HardIpInputData));
                var mergeList = new List<string>();
                int seqNum = curInterPoseInfo.Length > preMeasStr.Count
                    ? curInterPoseInfo.Length
                    : preMeasStr.Count;

                for (int i = 0; i < seqNum; i++)
                {
                    var seqMerge = new List<string>();
                    if (i < preMeasStr.Count && !string.IsNullOrEmpty(preMeasStr[i]))
                    {
                        seqMerge.Add(preMeasStr[i]);
                    }

                    if (i < curInterPoseInfo.Length && !string.IsNullOrEmpty(curInterPoseInfo[i]))
                    {
                        seqMerge.Add(curInterPoseInfo[i]);
                    }

                    if (seqMerge.Count > 0)
                    {
                        mergeList.Add(string.Join(">", seqMerge));
                    }
                }
                function.SetParamValue("InterPosePreMeas", string.Join("|", mergeList));
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private List<string> ProcessSocPreMeasToRf(string preMeas)
        {
            var result = new List<string>();
            foreach (string subPreMeas in preMeas.Split('|'))
            {
                var subResult = new List<string>();
                foreach (string inSubPreMeas in subPreMeas.Split(';'))
                {
                    string[] preMeasSegs = inSubPreMeas.Split(':');
                    if (preMeasSegs.Length >= 3)
                    {
                        subResult.Add(string.Format("Force{1}({0},{2})", preMeasSegs[0], preMeasSegs[1], preMeasSegs[2]));
                    }
                    else
                    {
                        subResult.Add(inSubPreMeas);
                    }
                }
                result.Add(string.Join(">", subResult));
            }
            return result;
        }


        private void SetMeasurement(HardIpPattern pattern)
        {
            var measPinsDic = pattern.MeasPins.GroupBy(p => p.SequenceIndex).ToDictionary(p => p.Key, p => p.ToList());
            Dictionary<int, List<MeasPin>> newGroups = ProcessMeasPinsOrder(measPinsDic);

            foreach (KeyValuePair<int, List<MeasPin>> measPins in newGroups)
            {
                if (measPins.Key < 1)
                {
                    continue;
                }

                _calclist.Add(GetCalcEquName(measPins.Value));
                _calcStorelist.Add(GetCalcStoreName(measPins.Value));
                _measStoreName.Add(GetStoreName(measPins.Value));
                _types.Add(GetTestType(measPins.Value));
            }
        }

        //Process non group Calc Meas Pins
        private Dictionary<int, List<MeasPin>> ProcessMeasPinsOrder(Dictionary<int, List<MeasPin>> groupsDictionary)
        {
            var result = new Dictionary<int, List<MeasPin>>();
            var curGroup = new KeyValuePair<int, List<MeasPin>>();
            var ignoreList = new List<string> { "BSTC", "BSTV", "VRFV", "VRFC" };
            foreach (KeyValuePair<int, List<MeasPin>> groupMeas in groupsDictionary)
            {
                if (groupMeas.Value.Exists(p => ignoreList.Exists(q => q.Equals(p.TestName))))
                {
                    continue;
                }

                if (groupMeas.Value.All(p => p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)))
                {
                    curGroup.Value.AddRange(groupMeas.Value);
                }
                else
                {
                    curGroup = groupMeas;
                    result.Add(curGroup.Key, curGroup.Value);
                }
            }
            return result;
        }


        private void SetInterPoseFunc(HardIpPattern pattern, Function function)
        {
            if (string.IsNullOrEmpty(pattern.RfInterPose))
            {
                return;
            }

            foreach (string interposeFunc in pattern.RfInterPose.Split(';'))
            {
                if (!interposeFunc.Contains(":"))
                {
                    continue;
                }

                string[] parainfo = interposeFunc.Split(':');
                string name = parainfo[0].Trim();
                string value = parainfo[1].Trim();
                int index = function.Parameters.Split(',').ToList().FindIndex(s => s.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (index != -1)
                {
                    function.ArgList[index] = value;
                }
            }
        }
        private string GetCalcEquName(List<MeasPin> measp)
        {
            IEnumerable<string> tmp = measp.Where(p => p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).Select(p => p.CalcEqn);
            var result = new List<string>();

            foreach (string item in tmp)
            {
                if (item != "")
                {
                    result.Add(item.Split(':')[1].Trim());
                }
                else
                {
                    result.Add(item);
                }
            }

            return string.Join(">", result);
        }

        private string GetCalcStoreName(List<MeasPin> measp)
        {
            IEnumerable<string> result = measp.Where(p => p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).Select(p => p.CusStr.Trim());


            return string.Join(">", result);
        }

        private string GetTestType(List<MeasPin> measp)
        {

            var pinsSetups = measp.Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase))
                .GroupBy(p => p.RfInstrumentSetup).ToDictionary(p => p.Key, p => p.ToList());
            var typeInseq = new List<string>();
            foreach (List<MeasPin> pinsSetup in pinsSetups.Values)
            {
                try
                {
                    var typeGrp = pinsSetup.GroupBy(p => p.InterPoseFunc).ToDictionary(p => p.Key, p => p.ToList());
                    typeInseq.Add(string.Join(",", typeGrp.Select(p => p.Key.Trim(';'))));
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                    throw;
                }

            }
            string type = string.Join("+", typeInseq);
            return type;
        }

        private string GetStoreName(List<MeasPin> measp)
        {
            var result = new List<string>();

            IEnumerable<MeasPin> pins = measp.Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase));
            foreach (MeasPin item in pins)
            {
                result.Add(item.CusStr.Trim());
            }

            return string.Join(">", result.Distinct());
        }

        private string AppendRegisterBitWidth(HardIpInfo info, string assignments)
        {
            var result = new List<string>();
            List<string> regList = info.SendBitName.Split('+').ToList();
            List<string> segList = info.SendBitStr.Split('+').ToList();
            foreach (string assignment in assignments.Replace(" ", "").Split(';'))
            {
                if (assignment.Contains("="))
                {
                    int regindex = regList.FindIndex(p => p.Equals(assignment.Split('=')[0], StringComparison.OrdinalIgnoreCase));
                    if (regindex != -1)
                    {
                        string segbit = segList[regindex].Split('_')[1];
                        result.Add(assignment + "|" + segbit);
                    }
                }
                else
                {
                    result.Add(assignment);
                }
            }
            return string.Join(";", result);
        }
        private string GetTrimType(string type)
        {
            if (Regex.IsMatch(type, "doall|domeasure|dotripoint|doquadpoint", RegexOptions.IgnoreCase))
            {
                if (type.Contains('@'))
                {
                    string[] arr = type.Split('@');
                    _sweepRange = arr[1];
                }
                return type.Split('@')[0];
            }

            if (type.IndexOf("binary", StringComparison.OrdinalIgnoreCase) > -1 ||
                type.IndexOf("custom", StringComparison.OrdinalIgnoreCase) > -1 ||
                type.IndexOf("nearbest", StringComparison.OrdinalIgnoreCase) > -1)
            {

                return type.Split('@')[0].ToUpper();
            }

            return "DOALL";
        }
    }
}
