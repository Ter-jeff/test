using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting;
using Automation.GenerateIgxl.Wireless.DVDC.WirelessConst;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Utility;

using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Wireless.DVDC.InstanceParameterSetting
{
    [ExcludeFromCodeCoverage]
    public class SetDvdcTrimValue : SetValueBase
    {
        private const string TransitionSlope = "L2H";
        private const string PutLimitsAllCodes = "TRUE";
        private const string DoApplyLevelsTiming = "TRUE";
        private string _doFinalVerification = "FALSE";
        private string _doTrimming = "TRUE";
        private string _slope = "Rising";
        private string _sweepRange = "";
        private readonly List<string> _2dsweepRange = new List<string>();

        private string _loadFile = "";

        private readonly List<string> _testSequenceList = new List<string>();
        private readonly List<string> _measPinList = new List<string>();
        private readonly List<string> _forceList = new List<string>();
        private readonly List<string> _irangeList = new List<string>();
        private readonly List<string> _vrangeList = new List<string>();
        private readonly List<string> _waitList = new List<string>();
        private readonly List<string> _calclist = new List<string>();

        public SetDvdcTrimValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            try
            {
                HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
                SetMeasurement(pattern);
                //patset
                function.ArgList[0] = pattern.Pattern.GetInstancePatternName();
                function.SetParamValue(WirelessConstData.TrimRegName, pattern.WirelessData.TrimFuseName.Split(',').Last());

                function.SetParamValue(WirelessConstData.TrimTarget, pattern.WirelessData.TrimTarget.Replace("&", "::"));

                GetTrimType(pattern.WirelessData.TrimType);
                function.SetParamValue(WirelessConstData.TrimType, "CUSTOM");
                //function.SetParamValue(WirelessConstData.TrimType, LocalSpecs.CurrentProject.ToLower().StartsWith("proxima") ? trimType : "CUSTOM");

                if (LocalSpecs.Options.Device == EnumDevice.RF &&
                    function.FunctionName.Equals(VbtFunctionLibShared.DvdcTrim, StringComparison.OrdinalIgnoreCase))
                {
                    function.SetParamValue(WirelessConstData.Stepsize, "-1");
                }
                else
                {
                    function.SetParamValue(WirelessConstData.Stepsize, GetStepSize(_sweepRange));
                }

                function.SetParamValue(WirelessConstData.TestSequence, string.Join(",", _testSequenceList));
                function.SetParamValue(WirelessConstData.MeasPins, GetParamFromList(_measPinList, "+"));
                function.SetParamValue(WirelessConstData.ForceConditions, RemoveUnit(GetParamFromList(_forceList)));
                function.SetParamValue(WirelessConstData.MeasIRange, RemoveUnit(GetParamFromList(_irangeList)));
                function.SetParamValue(WirelessConstData.MeasVRange, RemoveUnit(GetParamFromList(_vrangeList)));
                function.SetParamValue(WirelessConstData.MeasWaitTime, GetParamFromList(_waitList));
                function.SetParamValue(WirelessConstData.DigSrcpin, SearchInfo.GetSrcPin(info));
                function.SetParamValue(WirelessConstData.DigSrcEquation, info.SendBitName);
                function.SetParamValue(WirelessConstData.DigSrcAssignment,
                    AppendRegisterBitWidth(info, pattern.WirelessData.RegisterAssignment));
                function.SetParamValue(WirelessConstData.TrimCalcEqn, GetTrimCalcEq(pattern.WirelessData.TrimCalcEqn));
                function.SetParamValue(WirelessConstData.SweepRange, _sweepRange);
                function.SetParamValue(WirelessConstData.SweepRange2D, string.Join(",", _2dsweepRange.Where(x => !string.IsNullOrEmpty(x))));

                function.SetParamValue(WirelessConstData.Slope, _slope);
                function.SetParamValue(WirelessConstData.AdjustTrimCode, "'" + Regex.Replace(_sweepRange, "::", ","));

                function.SetParamValue(WirelessConstData.LoadFile, _loadFile);
                function.SetParamValue(WirelessConstData.TransitionSlope, TransitionSlope);
                function.SetParamValue(WirelessConstData.PutLimitsAllCodes, PutLimitsAllCodes);
                if (info.TrimRegName.Count > 1)//TwoDimensionTrimRegName
                {
                    List<string> dimensionTrimRegNameList = pattern.WirelessData.TrimFuseName.Split(',').ToList();
                    if (dimensionTrimRegNameList.Count > 1)
                    {
                        dimensionTrimRegNameList.RemoveRange(dimensionTrimRegNameList.Count - 1, 1);
                    }

                    function.SetParamValue("DimensionTrimRegName", string.Join(",", dimensionTrimRegNameList));
                }
                if (LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    _doTrimming = "TRUE";
                    _doFinalVerification = "TRUE";
                    function.SetParamValue(WirelessConstData.WriteFuncResult, "TRUE");
                    function.SetParamValue("DoApplyLevelsTiming", "True");
                    if (voltage.Equals("NV", StringComparison.CurrentCultureIgnoreCase))
                    {
                        function.SetParamValue(WirelessConstData.UpdateBstc2FuseDataBaseFlag, "-1");
                    }
                }
                else
                {
                    if (pattern.WirelessData.IsNeedPostBurn)
                    {
                        _doTrimming = "FALSE";
                        _doFinalVerification = "TRUE";
                        function.SetParamValue(WirelessConstData.WriteFuncResult, "TRUE");
                    }
                    else
                    {
                        _doTrimming = "TRUE";
                        _doFinalVerification = "FALSE";
                    }
                }

                function.SetParamValue(WirelessConstData.DoTrimming, _doTrimming);
                function.SetParamValue(WirelessConstData.DoFinalVerification, _doFinalVerification);

                function.SetParamValue(WirelessConstData.DigSrcBitWidth,
                    GetTrimRegBitWidth(info, pattern.WirelessData.RegisterAssignment));
                function.SetParamValue(WirelessConstData.ApplyLevelsTiming, DoApplyLevelsTiming);

                if (LocalSpecs.Options.Device == EnumDevice.LCD)
                {
                    SetInterPoseFunc(pattern, function, voltage);
                }
                else
                {
                    SetInterPoseFunc(pattern, function);
                }

                function.SetParamValue("DigCapPin", SearchInfo.GetMeasCPins(pattern, HardIpService.GetHardIpInfo(pattern)));
                //DigCap_DataWidth:  Get from "Cap Bit Str" in patInfo file. Like "wdr14_10+wdr23_10" ===> 10
                //function.ArgList[18] = Regex.Match(info.CapBitStr, @"^wdr\d+_(?<num>(\d+)).*").Groups["num"].ToString();
                function.SetParamValue("DigCapDataWidth", SearchInfo.GetDigDataWidth(info.CapBitStr));
                //DigCap_Sample_Size: Get from "Cap Bit" in patInfo file
                //function.ArgList[19] = info.CapBit.ToString();
                function.SetParamValue("DigCapSampleSize", info.CapBit.ToString("G15"));
                HardIpInfo hardIpInfo = HardIpService.GetHardIpInfo(pattern);
                function.SetParamValue("DigCapData_PostProcessing", SearchInfo.GetCusStrDigCapData(pattern, hardIpInfo));
                function.SetParamValue("CalcEquName", Filter(string.Join("|", _calclist)));
                function.SetParamValue("CalcStoreName", info.CalcStoreName);
                function.SetParamValue("MeasName", Filter(GetMeasName(pattern.MiscInfo)));
                function.SetParamValue("TrimloopBinOut", "True");
                MergeForceItemsToInterPosePreMeas(function, pattern);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
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
        private string AppendRegisterBitWidth(HardIpInfo info, string assignments)
        {
            var result = new List<string>();
            List<string> regList = info.SendBitName.Split('+').ToList();
            List<string> segList = info.SendBitStr.Split('+').ToList();
            string trimfuse = info.TrimRegName.LastOrDefault();
            foreach (string assignment in assignments.Replace(" ", "").Split(';'))
            {
                if (!string.IsNullOrEmpty(trimfuse) && assignment.Split('=')[0].Equals(trimfuse, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

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
            if (!string.IsNullOrEmpty(trimfuse))
            {
                result.Add(trimfuse.ToLower());
            }

            return string.Join(";", result);
        }

        private string GetMeasName(string misc)
        {
            string result = "";
            foreach (string item in misc.Split(';'))
            {
                if (Regex.IsMatch(item, "measname", RegexOptions.IgnoreCase) && item.Contains(":"))
                {
                    return item.Split(':')[1];
                }
            }
            return result;
        }

        private string GetStepSize(string sweepRange)
        {
            if (string.IsNullOrEmpty(sweepRange))
            {
                return "1";
            }

            if (sweepRange.Contains('('))
            {
                return "1";
            }

            string tempVal = sweepRange;
            string[] sweepVal = tempVal.Replace("::", "&").Split('&');
            if (decimal.Parse(sweepVal[0]) < decimal.Parse(sweepVal[1]))
            {
                return "1";
            }
            return "-1";
        }

        private string GetTrimRegBitWidth(HardIpInfo info, string digSrcAssignment)
        {
            string bitWidth = "0";
            try
            {
                if (!string.IsNullOrEmpty(digSrcAssignment))
                {
                    List<string> segs = info.SendBitStr.Split('+').ToList();
                    List<string> regs = info.SendBitName.Split('+').ToList();
                    string trimreg = info.TrimRegName.LastOrDefault();
                    if (!string.IsNullOrEmpty(trimreg))
                    {
                        int regIndex = regs.FindIndex(p => p.Equals(trimreg, StringComparison.OrdinalIgnoreCase));
                        if (regIndex != -1)
                        {
                            bitWidth = segs[regIndex].Split('_')[1];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return bitWidth;
        }

        private void SetMeasurement(HardIpPattern pattern)
        {
            var multiPin = new List<string>();
            var multiForce = new List<string>();
            var multiIrange = new List<string>();
            var multiVrange = new List<string>();
            var multiWait = new List<string>();
            var multiCalc = new List<string>();
            string trimMeas = pattern.WirelessData.TrimMeas;
            foreach (string meas in trimMeas.Split(';'))
            {
                if (meas == "")
                {
                    continue;
                }

                var testSeq = new List<string>();
                multiPin.Clear();
                multiForce.Clear();
                multiIrange.Clear();
                multiVrange.Clear();
                multiWait.Clear();
                multiCalc.Clear();
                foreach (string item in meas.Split('&'))
                {
                    string[] arr = item.Split(',');
                    string seq = SpiltMeasurement(arr, out string measPinName);
                    if (seq.Trim().Equals("calc", StringComparison.OrdinalIgnoreCase))
                    {
                        testSeq.Add(seq);
                        multiCalc.Add(arr[0].Split('@')[1]);
                    }
                    else if (seq == "C")
                    {
                        multiPin.Add(measPinName);
                    }
                    else
                    {
                        if (testSeq.Contains(seq))
                        {
                            continue;
                        }

                        multiPin.Add(measPinName);
                        testSeq.Add(seq.Replace("R1", "R"));
                        SetMeasValue(seq, arr, multiForce, multiIrange, multiVrange, multiWait);
                    }

                }
                _testSequenceList.Add(string.Join(">", testSeq));
                _calclist.Add(string.Join(">", multiCalc));
                _measPinList.Add(string.Join(",", multiPin));
                _forceList.Add(string.Join(",", multiForce));
                _irangeList.Add(string.Join(",", multiIrange));
                _vrangeList.Add(string.Join(",", multiVrange));
                _waitList.Add(string.Join(",", multiWait));
            }
        }

        private void SetMeasValue(string testSeq, IList<string> arrList, ICollection<string> multiForce, ICollection<string> multiIrange, ICollection<string> multiVrange, ICollection<string> multiWait)
        {
            switch (testSeq)
            {
                case "FI":
                case "FV":
                case "I":
                case "V":
                    if (arrList.Count > 1)
                    //multiForce.Add(arrList[1].Trim());
                    {
                        multiForce.Add(DataConvertor.ConvertUnits(RemoveUnit(arrList[1].Trim())));
                    }

                    if (arrList.Count > 2)
                    {
                        multiIrange.Add(arrList[2].Trim());
                    }

                    if (arrList.Count > 3)
                    {
                        multiVrange.Add(arrList[3].Trim());
                    }

                    if (arrList.Count > 4)
                    {
                        multiWait.Add(arrList[4].Trim());
                    }

                    break;
                case "R":
                case "R1":
                    if (arrList.Count > 1)
                    //multiForce.Add(arrList[1].Trim());
                    {
                        multiForce.Add(DataConvertor.ConvertUnits(RemoveUnit(arrList[1].Trim())));
                    }

                    if (arrList.Count > 2)
                    {
                        multiIrange.Add(arrList[2].Trim());
                    }

                    if (arrList.Count > 3)
                    {
                        multiWait.Add(arrList[3].Trim());
                    }

                    break;
                case "F":
                    multiWait.Add(arrList[1].Trim());
                    break;
                case "Vdm":
                case "Vdiff":
                    multiForce.Add(DataConvertor.ConvertUnits(arrList[1].Trim()));
                    multiVrange.Add(arrList[2].Trim());
                    multiWait.Add(arrList[3].Trim());
                    break;
                case "Idiff":
                    multiForce.Add(DataConvertor.ConvertUnits(arrList[1].Trim()));
                    multiIrange.Add(arrList[2].Trim());
                    multiWait.Add(arrList[3].Trim());
                    break;
                case "DutyCycle":
                    if (arrList.Count > 1)
                    {
                        multiForce.Add(DataConvertor.ConvertUnits(arrList[1].Trim()));
                    }

                    if (arrList.Count > 2)
                    {
                        multiIrange.Add(arrList[2].Trim());
                    }

                    if (arrList.Count > 3)
                    {
                        multiVrange.Add(arrList[3].Trim());
                    }

                    if (arrList.Count > 4)
                    {
                        multiWait.Add(arrList[4].Trim());
                    }

                    break;
            }
        }

        private string SpiltMeasurement(IList<string> arr, out string measPin)
        {
            measPin = "";
            string testSeq = "";

            if (arr[0].Contains('@'))
            {
                bool isCalc = arr[0].Split('@')[0].Equals("Calc", StringComparison.OrdinalIgnoreCase);
                if (isCalc)
                {
                    return arr[0].Split('@')[0];
                }

                string[] item = arr[0].Split('@');
                if (item[0].Trim().Equals("FVMI", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "I";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("FIMV", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "V";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("MeasF", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "F";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("MeasC", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "C";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("MeasR", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "R";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("MeasR1", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "R1";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("N", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "N";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("MeasVDM", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "Vdm";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("MeasVdiff", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "Vdiff";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("MeasIdiff", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "IDiff";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("Target_FIMV", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "V";
                    measPin = "Target_" + item[1].Trim();
                }
                if (item[0].Trim().Equals("Target_FVMI", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "I";
                    measPin = "Target_" + item[1].Trim();
                }
                if (item[0].Trim().Equals("FV", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "FV";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("FI", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "FI";
                    measPin = item[1].Trim();
                }
                if (item[0].Trim().Equals("DutyCycle", StringComparison.OrdinalIgnoreCase))
                {
                    testSeq = "DutyCycle";
                    measPin = item[1].Trim();
                }
            }
            return testSeq;
        }

        private string GetParamFromList(List<string> list, string delimit = "|")
        {
            if (list.Any(p => !string.IsNullOrEmpty(p)))
            {
                return string.Join(delimit, list);
            }

            return "";
        }

        private string GetTrimType(string types)
        {
            string result = "";
            List<string> trimSet = PreProcessTrimType(types);
            foreach (string type in trimSet)
            {
                _2dsweepRange.Add(_sweepRange);
                _sweepRange = "";
                if (Regex.IsMatch(type, "doall|domeasure|dotripoint|doquadpoint", RegexOptions.IgnoreCase))
                {
                    if (type.Contains('@'))
                    {
                        string[] arr = type.Split('@');

                        _sweepRange = arr[1];
                    }
                    result = type.Split('@')[0];
                }

                if (type.IndexOf("binary", StringComparison.OrdinalIgnoreCase) > -1 ||
                    type.IndexOf("custom", StringComparison.OrdinalIgnoreCase) > -1 ||
                    type.IndexOf("nearbest", StringComparison.OrdinalIgnoreCase) > -1)
                {

                    if (type.Contains('@'))
                    {
                        string regCus1 = @"(?<item>\([\d\,]+\))";
                        string[] arr = type.Split('@');
                        if (type.IndexOf("custom", StringComparison.OrdinalIgnoreCase) > -1)
                        {
                            if (Regex.IsMatch(arr[1], "loadfile", RegexOptions.IgnoreCase))
                            {
                                string regLoadFile = @"loadfile\s*\((?<path>.*)\)";
                                string path = Regex.Match(arr[1], regLoadFile, RegexOptions.IgnoreCase).Groups["path"].ToString();
                                _loadFile = path;
                            }
                            else
                            {
                                if (Regex.IsMatch(arr[1], regCus1, RegexOptions.IgnoreCase))
                                {
                                    var matches = Regex.Matches(arr[1], regCus1, RegexOptions.IgnoreCase).Cast<Match>().Select(p => p.Value).ToList();
                                    _sweepRange = string.Join(";", matches);
                                    _2dsweepRange.Add(_sweepRange);
                                }
                                else
                                {
                                    _sweepRange = arr[1];
                                }
                            }
                        }
                        else
                        {
                            _slope = arr[1];
                            if (_slope.Equals("rising", StringComparison.OrdinalIgnoreCase))
                            {
                                _slope = "Rising";
                            }
                            else if (_slope.Equals("falling", StringComparison.OrdinalIgnoreCase))
                            {
                                _slope = "Falling";
                            }

                            if (arr.Length > 2)
                            {
                                _sweepRange = arr[2];
                            }
                        }

                    }
                    result = type.Split('@')[0].ToUpper();
                }
            }
            if (string.IsNullOrEmpty(result))
            {
                result = "DOALL";
            }

            return result;
        }

        private string GetTrimCalcEq(string calceq)
        {
            string defaultCalceq = "algFindBestCode";
            if (!string.IsNullOrEmpty(calceq))
            {
                return calceq;
            }

            return defaultCalceq;
        }

        private void MergeForceItemsToInterPosePreMeas(Function function, HardIpPattern pattern)
        {
            try
            {
                List<string> curInterPoseInfo = function.GetParamValue("InterPosePreMeas").Split('|').ToList();
                List<string> preMeasStr = ProcessSocPreMeasToLcd(SearchInfo.GetPreMeas(pattern, HardIpInputData));
                var mergeList = new List<string>();
                int seqNum = curInterPoseInfo.Count > preMeasStr.Count
                    ? curInterPoseInfo.Count
                    : preMeasStr.Count;

                for (int i = 0; i < seqNum; i++)
                {
                    var seqMerge = new List<string>();
                    if (i < preMeasStr.Count && !string.IsNullOrEmpty(preMeasStr[i]))
                    {
                        seqMerge.Add(preMeasStr[i]);
                    }

                    if (i < curInterPoseInfo.Count && !string.IsNullOrEmpty(curInterPoseInfo[i]))
                    {
                        seqMerge.Add(curInterPoseInfo[i]);
                    }

                    if (seqMerge.Count > 0)
                    {
                        mergeList.Add(string.Join(">", seqMerge));
                    }
                }
                function.SetParamValue("InterPosePreMeas", string.Join("|", mergeList));

                curInterPoseInfo = new List<string> { function.GetParamValue("InterposePreInit") };
                curInterPoseInfo.AddRange(ProcessSocPreMeasToLcd(SearchInfo.GetPrePat(pattern), "&"));
                curInterPoseInfo.RemoveAll(string.IsNullOrEmpty);
                function.SetParamValue("InterposePreInit", string.Join("&", curInterPoseInfo));
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private List<string> ProcessSocPreMeasToLcd(string preMeas, string delimiter = ">")
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
                result.Add(string.Join(delimiter, subResult));
            }
            return result;
        }

        private new string Filter(string item)
        {
            if (item == "|" || item == "|,")
            {
                return "";
            }

            return item;
        }

        private List<string> PreProcessTrimType(string trimtype)
        {
            var result = new List<string>();
            var mainType = new List<string> { "doall", "domeasure", "dotripoint", "doquadpoint" };
            var otherType = new List<string> { "binary", "custom", "nearbest" };
            var setInfo = new List<string>();
            foreach (string info in trimtype.Split(';'))
            {
                if (mainType.Exists(info.ContainsIgnoreCase) ||
                    otherType.Exists(info.ContainsIgnoreCase))
                {
                    if (setInfo.Count > 0)
                    {
                        result.Add(string.Join(";", setInfo));
                    }

                    setInfo = new List<string>();
                }
                setInfo.Add(info);
            }
            if (setInfo.Count > 0)
            {
                result.Add(string.Join(";", setInfo));
            }

            return result;
        }
    }
}
