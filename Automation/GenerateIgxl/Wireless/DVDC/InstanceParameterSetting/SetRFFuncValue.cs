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
    public class SetRfFuncValue : SetValueBase
    {
        private readonly List<string> _calcStorelist = new List<string>();
        private readonly List<string> _measStoreName = new List<string>();
        private readonly List<string> _types = new List<string>();

        public SetRfFuncValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            /*pat	DsscSetup	RFInstSetup	testType ForceConditions	
             * InterposePreInit	InterposePrePat	InterposePostInit	Interpose_PreMeas	InterposePostMeas	InterposePreRst	
             * InterposePostPat	InterposePostRst	Meas_StoreName	CalcStoreName

             */
            HardIpService.GetHardIpInfo(pattern);
            string payload = pattern.Pattern.GetLastPayload();

            function.SetParamValue("Pat", payload);

            function.SetParamValue("DsscSetup", pattern.WirelessData.RegisterAssignment);

            string rfSetupName = SearchInfo.GetInstrumentInfo(pattern, "RFInstSetup");
            function.SetParamValue("RFInstSetup", rfSetupName);
            if (Regex.IsMatch(rfSetupName, "UWS|UWM|LXS|LXM|UPS|SwitchControl", RegexOptions.IgnoreCase))
            {
                function.SetParamValue("InterposePostRst", "setUWDisconnectPinFlag(TRUE)");
            }

            function.SetParamValue("DoApplyLevelsTiming", "TRUE");
            SetMeasurement(pattern);

            function.SetParamValue("SeqCnt", _types.Count.ToString());
            function.SetParamValue("TestType", string.Join("|", _types));
            function.SetParamValue("CalcEquName", GetCalcEquName(pattern.MeasPins));
            function.SetParamValue("CalcStoreName", string.Join("|", _calcStorelist));
            function.SetParamValue("MeasStoreName", string.Join("|", _measStoreName));
            function.SetParamValue(WirelessConstData.DoTrimming, "-1");
            function.SetParamValue(WirelessConstData.DoFinalVerification, "-1");
            if (Regex.IsMatch(pattern.MiscInfo, "isFW", RegexOptions.IgnoreCase))
            {
                function.SetParamValue(WirelessConstData.FwFlag, "TRUE");
            }

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

                //calclist.Add(GetCalcEquName(MeasPins.Value));
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
            foreach (KeyValuePair<int, List<MeasPin>> groupMeas in groupsDictionary)
            {
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
            if (measp.Count == 0)
            {
                return "";
            }

            var tmp = measp.Where(p => p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).Select(p => p.CalcEqn)
                .GroupBy(p => p.Split('(')[0]).ToDictionary(p => p.Key, p => p.ToList());
            var result = new List<string>();
            string regInfo = @"\((?<info>.*)\)";
            foreach (KeyValuePair<string, List<string>> item in tmp)
            {
                var info =
                    item.Value.Select(p => Regex.Match(p, regInfo, RegexOptions.IgnoreCase).Groups["info"].Value)
                        .ToList();
                result.Add($"{item.Key.Trim(':')}({string.Join(",", info)})");
            }
            int seqNum = measp.Max(p => p.SequenceIndex);
            string prefixedEmpty = string.Join("|", Enumerable.Repeat("", seqNum).ToList());


            return prefixedEmpty + string.Join(";", result);
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
                    var types = typeGrp.Select(p => p.Key.Trim(';')).ToList();
                    for (int i = 0; i < types.Count; i++)
                    {
                        if (string.IsNullOrEmpty(types[i]))
                        {
                            types[i] = "N";
                        }
                    }
                    typeInseq.Add(string.Join(",", types));
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
    }
}
