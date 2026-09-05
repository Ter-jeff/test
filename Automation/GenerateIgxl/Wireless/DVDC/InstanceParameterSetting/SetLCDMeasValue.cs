using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting;
using Automation.GenerateIgxl.PreAction.GenDSSCSetup;
using Automation.GenerateIgxl.Wireless.DVDC.WirelessConst;
using Automation.InputManager.Data;
using Automation.Utility.HardIP;

using LogLib.Utility;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Wireless.DVDC.InstanceParameterSetting
{
    [ExcludeFromCodeCoverage]
    public class SetLcdMeasValue : SetValueBase
    {
        private const string DoApplyLevelsTiming = "TRUE";

        private static readonly Regex _regRunPats = new Regex(
            @"Run(?:Init|Payload)Pat\(\s*(?<arg>[^)]*)\)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private readonly List<string> _testSequenceList = new List<string>();
        private readonly List<string> _measPinList = new List<string>();
        private readonly List<string> _forceList = new List<string>();
        private readonly List<string> _irangeList = new List<string>();
        private readonly List<string> _vrangeList = new List<string>();
        private readonly List<string> _waitList = new List<string>();
        private readonly List<string> _calclist = new List<string>();
        private readonly List<string> _calcStorelist = new List<string>();
        private readonly List<string> _measStoreName = new List<string>();
        private readonly List<string> _measName = new List<string>();

        public SetLcdMeasValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            try
            {
                SetMeasurement(pattern);
                HardIpService.GetHardIpInfo(pattern);
                string forceflag = GetForceFlagInfo(pattern.MiscInfo);
                string measpins = Filter(string.Join("+", _measPinList));
                string keepForcePostPat = GetIntPostPatFromFlag(forceflag, measpins);

                List<string> allPats = new List<string>();
                if (pattern.MiscInfoDict.ContainsKey("InterposePrePat"))
                {
                    allPats.AddRange(ExtractRunPats(pattern.MiscInfoDict["InterposePrePat"]));
                }
                allPats.AddRange(pattern.Pattern.GetInstancePatternName().Split(';'));
                if (pattern.MiscInfoDict.ContainsKey("InterposePostPat"))
                {
                    allPats.AddRange(ExtractRunPats(pattern.MiscInfoDict["InterposePostPat"]));
                }

                //patset
                function.ArgList[0] = pattern.Pattern.GetInstancePatternName();

                function.SetParamValue(WirelessConstData.TestSequence, Filter(string.Join(",", _testSequenceList)));
                function.SetParamValue(WirelessConstData.ForceConditions, Filter(string.Join("|", _forceList)));
                function.SetParamValue(WirelessConstData.MeasIRange, Filter(string.Join("|", _irangeList)));
                function.SetParamValue(WirelessConstData.MeasVRange, Filter(string.Join("|", _vrangeList)));
                function.SetParamValue(WirelessConstData.MeasWaitTime, Filter(string.Join("|", _waitList)));

                SetInterPoseFunc(pattern, function, voltage, keepForcePostPat); //All interpose should be execute this function to generate external table
                MergeForceItemsToInterPosePreMeas(function);

                function.SetParamValue(WirelessConstData.CalcEquName, Filter(string.Join("|", _calclist)));
                function.SetParamValue(WirelessConstData.CalcStoreName, Filter(string.Join("|", _calcStorelist)));

                function.SetParamValue(WirelessConstData.MeasPins, measpins);
                function.SetParamValue(WirelessConstData.MeasStoreName, Filter(string.Join("|", _measStoreName)));
                function.SetParamValue(WirelessConstData.MeasName, Filter(string.Join("|", _measName)));
                function.SetParamValue(WirelessConstData.MeasLimit, GetMeasLimit(pattern.MeasPins));

                function.SetParamValue(WirelessConstData.PinsKeepForceFlag, JudgeKeepForce(pattern));
                function.SetParamValue(WirelessConstData.ApplyLevelsTimings, DoApplyLevelsTiming);
                function.SetParamValue(WirelessConstData.PinsKeepForceFlag, forceflag);

                string dsscSheetName = "DSSCSetupSheet_" + CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
                string dsscSetupKey = pattern.RegAssignName + "_DSSCSetup";

                var dsscSetupInst = new DsscSetupSheet(dsscSetupKey, new List<DsscItem>());

                foreach (string patternname in allPats)
                {
                    HardIpInfo info = HardIpService.GetHardIpInfo(patternname);

                    dsscSetupInst.Items.Add(new DsscItem(
                        info.Payload == "" ? "(X)" + patternname : info.Payload,
                        info.SendBitName,
                        info.DigSrcAssignment,
                        info.SendPinName,
                        info.SendBitStr,
                        info.CapPinName,
                        info.CapBit,
                        "",
                        info.FullPattern
                    ));
                }

                DsscSetupMain.Save(dsscSheetName, dsscSetupInst);

                function.SetParamValue(WirelessConstData.DSSCSetup, dsscSetupKey);

            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        public static List<string> ExtractRunPats(string input)
        {
            List<string> result = new List<string>();
            if (string.IsNullOrWhiteSpace(input))
            {
                return result;
            }

            MatchCollection matches = _regRunPats.Matches(input);

            foreach (Match m in matches)
            {
                string rawArg = m.Groups["arg"].Value;
                if (string.IsNullOrWhiteSpace(rawArg))
                {
                    continue;
                }

                string[] parts = rawArg.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string p in parts)
                {
                    result.Add(p.Trim());
                }
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

        private string GetIntPostPatFromFlag(string forceFlag, string measPins)
        {
            return forceFlag.Equals("0") || forceFlag.Equals("") ? "" : "Disconnectpins(" + measPins.Replace("+", ",").Replace("::", ",") + ")";
        }

        private string GetForceFlagInfo(string miscinfo)
        {
            return Regex.Match(miscinfo, "pinsKeepForceFlag:(?<value>.+);", RegexOptions.IgnoreCase).Groups["value"].ToString();
        }

        private string GetMeasLimit(IEnumerable<MeasPin> measPins)
        {

            var limitList = new List<string>();
            foreach (MeasPin pins in measPins)
            {
                if (!pins.MeasLimitsN.Any())
                {
                    continue; //20220921LCD currently only NV case
                }

                var limitNv =
                    pins.MeasLimitsN.Where(x => !string.IsNullOrEmpty(x.HiLimit) && !string.IsNullOrEmpty(x.LoLimit)).ToList();

                var pinLimitList = new List<string>();
                foreach (MeasLimit limit in limitNv)
                {
                    string unit = Regex.Match(limit.HiLimit, "(?<unit>(A|V|Hz|ohm).*)", RegexOptions.IgnoreCase).Groups["unit"].ToString();
                    string limitStr =
                        $"{limit.JobName}={DataConvertor.ConvertUnits(limit.HiLimit)},{DataConvertor.ConvertUnits(limit.LoLimit)},{unit}";
                    pinLimitList.Add(limitStr);
                }
                limitList.Add(string.Join(";", pinLimitList));
            }

            return Filter(string.Join("|", limitList));
        }

        private string JudgeKeepForce(HardIpPattern pattern)
        {
            string result = "False";
            try
            {
                if (
                    pattern.MiscInfo.Split(';')
                        .ToList()
                        .Exists(p => Regex.IsMatch(p.Split(':')[0], "KeepForce", RegexOptions.IgnoreCase)))
                {
                    string keepForceInfo =
                        pattern.MiscInfo.Split(';')
                            .FirstOrDefault(p => Regex.IsMatch(p.Split(':')[0], "KeepForce", RegexOptions.IgnoreCase));
                    result = keepForceInfo.Split(':')[1];

                }
                else if (pattern.MeasPins.Any(k => k.MiscInfo.Split(';').ToList()
                    .Exists(p => Regex.IsMatch(p.Split(':')[0], "KeepForce", RegexOptions.IgnoreCase))))
                {
                    string keepForceSyntaxPin = pattern.MeasPins.FirstOrDefault(k => k.MiscInfo.Split(';').ToList()
                        .Exists(p => Regex.IsMatch(p.Split(':')[0], "KeepForce", RegexOptions.IgnoreCase)))
                        .MiscInfo.Split(';')
                        .FirstOrDefault(p => Regex.IsMatch(p.Split(':')[0], "KeepForce", RegexOptions.IgnoreCase));
                    result = keepForceSyntaxPin.Split(':')[1];

                }
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
            return result;
        }

        private void MergeForceItemsToInterPosePreMeas(Function function)
        {
            try
            {
                List<string> curInterPoseInfo = function.GetParamValue("InterPosePreMeas").Split('|').ToList();
                List<string> preMeasStr = ProcessSocPreMeasToLcd();
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
                        mergeList.Add(string.Join("&", seqMerge));
                    }
                }
                function.SetParamValue("InterPosePreMeas", string.Join("|", mergeList));

                curInterPoseInfo = new List<string> { function.GetParamValue("InterposePreInit") };
                curInterPoseInfo.AddRange(ProcessSocPreMeasToLcd());
                curInterPoseInfo.RemoveAll(string.IsNullOrEmpty);
                function.SetParamValue("InterposePreInit", string.Join("&", curInterPoseInfo));
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
            }
        }

        private List<string> ProcessSocPreMeasToLcd()
        {
            return new List<string>();
        }

        private string GetMeasName(IEnumerable<MeasPin> measp)
        {
            var result = new List<string>();

            IEnumerable<MeasPin> pins = measp.Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase));
            foreach (MeasPin item in pins)
            {
                result.Add(item.TestName);
            }

            return string.Join(">", result);
        }

        private string GetStoreName(IEnumerable<MeasPin> measp)
        {
            var result = new List<string>();

            IEnumerable<MeasPin> pins = measp.Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase));
            foreach (MeasPin item in pins)
            {
                result.Add(item.CusStr);
            }

            if (result.Any(p => !string.IsNullOrEmpty(p)))
            {
                return string.Join(">", result);
            }

            return null;
        }
        private string GetCalcEquName(IEnumerable<MeasPin> measp)
        {
            IEnumerable<string> tmp = measp.Where(p => p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).Select(p => p.CalcEqn);
            var result = new List<string>();

            foreach (string item in tmp)
            {
                if (item != "")
                {
                    result.Add(item.Split(':')[1]);
                }
                else
                {
                    result.Add(item);
                }
            }

            return string.Join(">", result);
        }

        private string GetCalcStoreName(IEnumerable<MeasPin> measp)
        {
            IEnumerable<string> result = measp.Where(p => p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).Select(p => p.CusStr);


            return string.Join(">", result);
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

                _testSequenceList.Add(GetTestSequence(measPins.Value.OrderBy(x => x.SequenceIndex).ToList()));
                _calclist.Add(GetCalcEquName(measPins.Value));
                _calcStorelist.Add(GetCalcStoreName(measPins.Value));
                _measPinList.Add(GetMeasPin(measPins.Value));
                if (!measPins.Value.Exists(x => x.MeasType.Equals(MeasType.MeasN, StringComparison.CurrentCultureIgnoreCase)))
                {
                    _measName.Add(GetMeasName(measPins.Value));
                }

                _measStoreName.Add(GetStoreName(measPins.Value));
                _forceList.Add(string.Join(",", GetForceValue(measPins.Value
                    .Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).ToList())));
                _irangeList.Add(GetMeasInfo(measPins.Value
                    .Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).ToList()
                    , "MeasI_Range", "999"));
                _vrangeList.Add(GetMeasInfo(measPins.Value
                    .Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).ToList(), "MeasVRange", "999"));
                _waitList.Add(GetMeasInfo(measPins.Value
                    .Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).ToList(), "MeasWaitTime", "0"));
            }
        }

        //Process non group Calc Meas Pins
        private Dictionary<int, List<MeasPin>> ProcessMeasPinsOrder(Dictionary<int, List<MeasPin>> groupsDictionary)
        {
            var result = new Dictionary<int, List<MeasPin>>();
            var curGroup = new KeyValuePair<int, List<MeasPin>>(0, new List<MeasPin>());
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

        private string GetTestSequence(List<MeasPin> pins)
        {
            var seq = new List<string>();
            if (pins.Exists(p => p.MeasType.Equals(MeasType.MeasN, StringComparison.OrdinalIgnoreCase)))
            {
                string type = "";
                var forceList = pins.SelectMany(p => p.ForceConditions).ToList();
                if (forceList.Any())
                {
                    type = forceList.SelectMany(p => p.ForcePins).First().ForceType;
                }
                else if (pins.Any(p => !string.IsNullOrEmpty(DeriveMiscInfo(p.MiscInfo, "MeasWaitTime"))))
                {
                    return "MeasWait";
                }

                if (!string.IsNullOrEmpty(type))
                {
                    seq.Add(type.StartsWith("P") ? type : "F" + type);
                }
                else
                {
                    seq.Add("N");
                }
            }
            else
            {
                seq.Add(Regex.Replace(pins.First().MeasType, "meas", "", RegexOptions.IgnoreCase));
            }
            var calcEqn = pins.Where(p => !string.IsNullOrEmpty(p.CalcEqn)).Select(p => p.MeasType).ToList();
            if (calcEqn.Any())
            {
                seq.AddRange(calcEqn);
            }

            return string.Join(">", seq);
        }

        private IEnumerable<string> GetForceValue(IEnumerable<MeasPin> pins)
        {
            var valueList = new List<string>();
            foreach (MeasPin measPin in pins)
            {
                var force = measPin.ForceConditions.SelectMany(p => p.ForcePins).ToList();
                string value = "0";
                //value = !force.Any(p => p.PinName.Equals(measPin.PinName, StringComparison.OrdinalIgnoreCase)) ? "0" : force.First(p => p.PinName.Equals(measPin.PinName, StringComparison.OrdinalIgnoreCase)).ForceValue;

                if (force.Any(p => p.PinName.Equals(measPin.PinName, StringComparison.OrdinalIgnoreCase)))
                {
                    value = force.First(p => p.PinName.Equals(measPin.PinName, StringComparison.OrdinalIgnoreCase)).ForceValue;
                }
                else if (force.Any())
                {
                    value = force.First().ForceValue;
                }

                valueList.Add(DataConvertor.ConvertUnits(RemoveUnit(value)));
            }

            return valueList;
        }

        private string GetMeasPin(List<MeasPin> pins)
        {
            if (pins.Exists(p => p.MeasType.Equals(MeasType.MeasN, StringComparison.OrdinalIgnoreCase)))
            {
                var forceList = pins.SelectMany(p => p.ForceConditions).ToList();
                if (forceList.Any())
                {
                    return string.Join(",", forceList.SelectMany(p => p.ForcePins).Select(p => p.PinName).ToList());
                }

                return "";
            }

            return string.Join(",", pins.Where(p => !p.MeasType.Equals(MeasType.MeasCalc, StringComparison.OrdinalIgnoreCase)).Select(p => p.PinName));
        }

        private string GetMeasInfo(IEnumerable<MeasPin> pins, string selectType, string defaultValue)
        {
            var infos = pins.Select(p => DeriveMiscInfo(p.MiscInfo, selectType)).ToList();
            if (infos.All(string.IsNullOrEmpty))
            {
                return defaultValue;
            }

            var valueList = new List<string>();
            foreach (string info in infos)
            {
                if (string.IsNullOrEmpty(info))
                {
                    valueList.Add(defaultValue);
                }
                else
                {
                    valueList.Add(DataConvertor.ConvertUnits(RemoveUnit(info)));
                }
            }
            return string.Join(",", valueList);
        }

        private string DeriveMiscInfo(string miscinfo, string item)
        {
            foreach (string miscitem in miscinfo.Split(';'))
            {
                if (!miscitem.Contains(":"))
                {
                    continue;
                }

                if (miscitem.Split(':')[0].Equals(item, StringComparison.OrdinalIgnoreCase))
                {
                    return miscitem.Split(':')[1];
                }
            }
            return "";
        }


    }
}
