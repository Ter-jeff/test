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
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using LogLib.Utility;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    public class ForceClass
    {
        private string _forceStr = "";

        public bool IsVtShmoo { get; set; }
        public bool IsShmooInForce { get; set; }
        public bool IsShmooInProdInst { get; set; }
        public bool IsShmooInProdFlow { get; set; }
        public bool IsShmooInCharInst { get; set; }
        public bool IsShmooInCharFlow { get; set; }
        public bool IsCz2InstName { get; set; }
        public string ForceCondition
        {
            get { return _forceStr; }
            set { _forceStr = CorrectTypo(value); }
        }

        //Dc setting eg: Level:XXX
        private const string RegLevelSetting = @"^(Level:|Levels:)(?<level>[\w]+)";
        //Ac setting eg: AC:XXX:XXX
        private const string RegAcSetting = @"^AC:[\w|&|,]+:[\w]+";
        //Ac Category eg: AC:XXX
        private const string RegAcCategory = @"^AC:[\w]+";
        //Ac selector eg: ACSelector:NV:XXX
        private const string RegAcSelector = @"ACSelector:[\w|&]+:[\w]+";
        //DC Category eg: DC:XXX
        private const string RegDcCategory = @"^DC:[\w]+";
        //Dc selector eg: DCSelector:NV:XXX
        private const string RegDcSelector = @"DCSelector:[\w|&]+:[\w]+";
        //TimeSets eg: TimeSets:XXX
        private const string RegTimeSet = @"^TimeSets:[\w]+";

        public ForceClass()
        {
            IsVtShmoo = false;
            IsShmooInForce = false;
            IsShmooInProdInst = true;
            IsShmooInProdFlow = true;
            IsShmooInCharInst = false;
            IsShmooInCharFlow = false;
            IsCz2InstName = false;
            ForceCondition = "";
        }

        public ForceClass(ForceClass other)
        {
            if (other == null)
            {
                return;
            }

            IsVtShmoo = other.IsVtShmoo;
            IsShmooInForce = other.IsShmooInForce;
            IsShmooInProdInst = other.IsShmooInProdInst;
            IsShmooInProdFlow = other.IsShmooInProdFlow;
            IsShmooInCharInst = other.IsShmooInCharInst;
            IsShmooInCharFlow = other.IsShmooInCharFlow;
            IsCz2InstName = other.IsCz2InstName;

            _forceStr = other._forceStr;
        }

        public ForceClass Copy()
        {
            return new ForceClass(this);
        }

        public HardipCharSetup GetShmoo(TestPlanSheet planSheet, HardIpPattern pattern, string pattForceCondition, string subBlockName)
        {
            try
            {
                bool isSweep = false;
                var result = new HardipCharSetup();
                string originalParameterName = "";
                foreach (string forceCondition in pattForceCondition.Split(';'))
                {

                    Match shmooMatch = HardIpConstData.RegVtShmoo.Match(forceCondition);
                    if (!shmooMatch.Success)
                    {
                        shmooMatch = HardIpConstData.RegShmoo.Match(forceCondition);
                    }

                    if (!shmooMatch.Success)
                    {
                        continue;
                    }

                    string shmoostr = shmooMatch.Groups["ShmooStr"].Value;

                    //Eg. Sweep(USBDPDM:vih:0.6,1.2,0.005:retest:jump,6)
                    string[] strArr = shmoostr.Replace("::", "&&").Split(':');//change differential pingroup type format to avoid split rule 

                    var oneShmooStep = new CharStep(pattern.Pattern.GetLastPayload(), pattern.PatternType)
                    {
                        Mode = Regex.IsMatch(forceCondition, "yshmoo|sweepy", RegexOptions.IgnoreCase) ? CharSetupConst.ModeYShmoo : CharSetupConst.ModeXShmoo,
                        VoltageType = GetShmooInstanceVoltage(result, forceCondition)
                    };

                    if (strArr.Length < 3)
                    {
                        return new HardipCharSetup();
                    }

                    string pinName = SearchInfo.GenDiffGroupName(strArr[0].Trim().Replace("&&", "::"), true);
                    string firstWord = strArr[1].Split(',')[0];

                    originalParameterName = strArr[1];

                    oneShmooStep.ApplyToPinExecMode = "Simultaneous";
                    oneShmooStep.StepName = pinName.Replace(",", "_");
                    oneShmooStep.ParameterName = HardipCharSetup.GetShmooParameterName(firstWord);
                    oneShmooStep.ParameterType = CharSetupSingleton.Instance().GetShmooParameterType(firstWord);
                    oneShmooStep.ApplyToTimeSets = HardipCharSetup.GetShmooTimeSets(strArr[1]);
                    ResolveParameterTypeAndName(oneShmooStep, pattern, firstWord);

                    if (Regex.IsMatch(oneShmooStep.ParameterType, CharSetupConst.ParameterTypeAcSpec, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(oneShmooStep.ParameterType, CharSetupConst.ParameterTypeDcSpec, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(oneShmooStep.ParameterType, CharSetupConst.ParameterTypeGlobalSpec, RegexOptions.IgnoreCase))
                    {
                        oneShmooStep.ApplyToPins = "";
                    }
                    else
                    {
                        oneShmooStep.ApplyToPins = DataConvertor.ConvertToNetName(strArr[0], TestPlanStatic.PowerMergeSheet.PowerMerge);
                    }

                    oneShmooStep.StepName = DataConvertor.ConvertToNetName(strArr[0], TestPlanStatic.PowerMergeSheet.PowerMerge).Replace(",", "_") + "_" + oneShmooStep.ParameterName;

                    string[] fromArr = strArr[2].Split(',');
                    if (fromArr.Length == 3)
                    {
                        oneShmooStep.RangeCalcField = CharSetupConst.RangeCalcFieldSteps;
                        oneShmooStep.RangeFrom = DataConvertor.ConvertUnits(fromArr[0].Trim());

                        oneShmooStep.RangeTo = DataConvertor.ConvertUnits(fromArr[1].Trim());
                        oneShmooStep.RangeStepSize = DataConvertor.ConvertUnits(fromArr[2].Trim());
                    }

                    ApplyTestMethodAndAlgorithm(oneShmooStep, result, strArr, forceCondition);

                    List<ProtocolAwarePin> nwirePins = NwireSingleton.Instance().SettingInfo.NwirePins;
                    if (nwirePins.Any(x => x.OutClk.ContainsIgnoreCase(pinName.ToUpper()) || x.OutClkDiff.ContainsIgnoreCase(pinName.ToUpper())))
                    {
                        ApplyNwireClockSettings(oneShmooStep, nwirePins, pinName);
                    }

                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharFuncNamePrintShmooInfoMain, "hardip", true);
                    if (function.Type == ".NET")
                    {
                        oneShmooStep.PostStep = function.FullFunctionName;
                    }
                    else
                    {
                        oneShmooStep.PostStep = CharSetupConst.PostStepFunctionPrintShmooInfo;
                    }

                    oneShmooStep.PostStepArguments = string.IsNullOrEmpty(oneShmooStep.ApplyToPins)
                        ? "CorePower"
                        : "CorePower," + oneShmooStep.ApplyToPins;

                    if (Regex.IsMatch(forceCondition, "sweep|sweepy", RegexOptions.IgnoreCase))
                    {
                        oneShmooStep.OutputSuspendDatalog = "FALSE";
                        isSweep = true;
                    }
                    else if (Regex.IsMatch(forceCondition, "xshmoo|yshmoo", RegexOptions.IgnoreCase))
                    {
                        oneShmooStep.OutputSuspendDatalog = "TRUE";
                        isSweep = false;
                    }

                    oneShmooStep.OutputFormat = "Enhanced";
                    oneShmooStep.OutputDestinationsTextFile = "Disable";
                    oneShmooStep.OutputDestinationsSheet = "Disable";
                    oneShmooStep.OutputDestinationsDatalog = "Enable";
                    oneShmooStep.OutputDestinationsImmediateWin = "Disable";
                    oneShmooStep.OutputDestinationsOutputWin = "Disable";

                    result.AddStep(oneShmooStep);
                }

                if (result.CharSteps.Any())
                {
                    string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
                    if (string.IsNullOrEmpty(subBlockName))
                    {
                        subBlockName = "SubBlock";
                    }

                    result.SetupName = HardIpConstData.PrefixShmooSetupName + blockName + "_" +
                                       CommonGenerator.GetSubBlockNameWithoutMinus(subBlockName) +
                                       pattern.PatternIndexFlag;
                    result.TestNameInFlow = HardipCharSetup.GetShmooName(planSheet, pattern, result, subBlockName, isSweep, originalParameterName);
                }
                return result;
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                return null;
            }
        }

        private void ResolveParameterTypeAndName(CharStep oneShmooStep, HardIpPattern pattern, string firstWord)
        {
            if (string.IsNullOrEmpty(oneShmooStep.ParameterType) || pattern.ForceCondition.IsVtShmoo)
            {
                if (string.IsNullOrEmpty(oneShmooStep.ParameterType))
                {
                    oneShmooStep.ParameterType = CharSetupConst.ParameterTypeLevel;
                    oneShmooStep.ParameterName = "Vmain";
                }
                else
                {
                    oneShmooStep.ParameterType = CharSetupConst.ParameterTypeGlobalSpec;
                    oneShmooStep.ParameterName = "Shmoo_Glb";
                }

                if (firstWord.Split('$').Length == 2)
                {
                    oneShmooStep.ParameterName = $"Shmoo_{firstWord.Split('$')[1]}_Glb";
                    TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value?.AddRow(new GlobalSpec(oneShmooStep.ParameterName, "0"));
                }
            }
        }

        private void ApplyTestMethodAndAlgorithm(CharStep oneShmooStep, HardipCharSetup result, string[] strArr, string forceCondition)
        {
            if (strArr.Length == 3)
            {
                if (oneShmooStep.ParameterType == CharSetupConst.ParameterTypeAcSpec)
                {
                    //Timing Sweep
                    result.TestMethod = CharSetupConst.TestMethodRetest;
                    oneShmooStep.AlgorithmName = CharSetupConst.AlgorithmNameLinear;
                }
                else
                {
                    if (Regex.IsMatch(forceCondition, "sweepy|yshmoo", RegexOptions.IgnoreCase))
                    {
                        result.TestMethod = CharSetupConst.TestMethodRetest;
                        oneShmooStep.AlgorithmName = CharSetupConst.AlgorithmNameLinear;
                    }
                    else
                    {
                        //Default Sweep
                        result.TestMethod = CharSetupConst.TestMethodRetest;
                        oneShmooStep.AlgorithmName = CharSetupConst.AlgorithmNameJump;
                        oneShmooStep.AlgorithmArguments = "6";
                    }
                }
            }

            if (strArr.Length == 4)
            {
                result.TestMethod = CharSetupConst.TestMethod.ContainsKey(strArr[3]) ? CharSetupConst.TestMethod[strArr[3]] : "";
            }
            else if (strArr.Length == 5)
            {
                result.TestMethod = CharSetupConst.TestMethod.ContainsKey(strArr[3]) ? CharSetupConst.TestMethod[strArr[3]] : "";
                string[] algorithmArr = strArr[4].Split(',');

                if (algorithmArr.Length == 1)
                {
                    oneShmooStep.AlgorithmName = algorithmArr[0];
                }
                else if (algorithmArr.Length == 2)
                {
                    oneShmooStep.AlgorithmName = algorithmArr[0];
                    oneShmooStep.AlgorithmArguments = algorithmArr[1];
                }
            }
            else if (strArr.Length == 6)
            {
                result.TestMethod = CharSetupConst.TestMethod.ContainsKey(strArr[3]) ? CharSetupConst.TestMethod[strArr[3]] : "";

                string[] algorithmArr = strArr[4].Split(',');
                if (algorithmArr.Length == 1)
                {
                    oneShmooStep.AlgorithmName = algorithmArr[0];
                }
                else if (algorithmArr.Length == 2)
                {
                    oneShmooStep.AlgorithmName = algorithmArr[0];
                    oneShmooStep.AlgorithmArguments = algorithmArr[1];
                }

                string[] parameterArr = strArr[5].Split(',');
                if (CharSetupConst.ParameterType.TryGetValue(parameterArr[0], out string paratypevalue))
                {
                    parameterArr[0] = paratypevalue;
                }
                if (CharSetupConst.ParameterName.TryGetValue(parameterArr[1], out string paranamevalue))
                {
                    parameterArr[1] = paranamevalue;
                }
                oneShmooStep.ParameterType = parameterArr[0];
                oneShmooStep.ParameterName = parameterArr[1];
            }
        }

        private void ApplyNwireClockSettings(CharStep oneShmooStep, List<ProtocolAwarePin> nwirePins, string pinName)
        {
            List<ProtocolAwarePin> pin = nwirePins.Where(x => x.OutClk.ContainsIgnoreCase(pinName.ToUpper()) || x.OutClkDiff.ContainsIgnoreCase(pinName.ToUpper())).ToList();
            string portName = pin.FirstOrDefault().CreatePortName(EnumEquipment.UltraFlex);
            string freqVarName = pin.FirstOrDefault().CreateFreqVarName();
            oneShmooStep.PrePoint = "freerunclk_set_XY";
            if (oneShmooStep.Mode == CharSetupConst.ModeYShmoo)
            {
                oneShmooStep.PrePointArguments = string.Format("Y," + portName + "," + freqVarName);
            }
            else
            {
                oneShmooStep.PrePointArguments = string.Format("X," + portName + "," + freqVarName);
            }

            oneShmooStep.PostPoint = "freerunclk_stop";
            oneShmooStep.PostPointArguments = string.Format(portName);
            oneShmooStep.StepName = freqVarName;
            oneShmooStep.ParameterType = CharSetupConst.ParameterTypeAcSpec;
            oneShmooStep.ParameterName = freqVarName;
            oneShmooStep.ApplyToPins = "";
        }

        public string GetShmooInstanceVoltage(HardipCharSetup shmooSetup, string forceCondition)
        {
            if (forceCondition.StartsWith("NV@", StringComparison.OrdinalIgnoreCase))
            {
                shmooSetup.IsSplitByVoltage = true;
                return "NV";
            }

            if (forceCondition.StartsWith("HV@", StringComparison.OrdinalIgnoreCase))
            {
                shmooSetup.IsSplitByVoltage = true;
                return "HV";
            }

            if (forceCondition.StartsWith("LV@", StringComparison.OrdinalIgnoreCase))
            {
                shmooSetup.IsSplitByVoltage = true;
                return "LV";
            }

            return "";
        }

        public string GetLevelSetting()
        {
            foreach (string force in ForceCondition.Split(';'))
            {
                if (Regex.IsMatch(force, RegLevelSetting, RegexOptions.IgnoreCase))
                {
                    return force.Split(':')[1];
                }
            }
            return "";
        }

        public string GetAcSetting()
        {
            string acSettings = "";
            foreach (string force in ForceCondition.Split(';'))
            {
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase))
                {
                    if (IsAcSpecPin(Regex.Replace(force, "::", "&").Split(':')[1]))
                    {
                        acSettings += force + ";";
                    }
                }
            }
            return acSettings.Trim(';');
        }

        private bool IsAcSpecPin(string pinName)
        {
            return Regex.IsMatch(pinName, "^(TCK|ShiftIn)$") ||
                NwireSingleton.Instance().SettingInfo.NwirePins.Find(s => s.OutClk.Equals(pinName, StringComparison.OrdinalIgnoreCase)) != null;
        }

        public string GetAcSelector()
        {
            string acAcSelector = "";
            foreach (string force in ForceCondition.Split(';'))
            {
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSelector, RegexOptions.IgnoreCase))
                {
                    acAcSelector += force + ";";
                }
            }
            return acAcSelector.Trim(';');
        }

        public string GetDcCategory()
        {
            string dcSettings = "";
            foreach (string force in ForceCondition.Split(';'))
            {
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcCategory, RegexOptions.IgnoreCase))
                {
                    dcSettings += force + ";";
                }
            }
            return dcSettings.Trim(';');
        }

        public string GetAcCategory()
        {
            string acSettings = "";
            foreach (string force in ForceCondition.Split(';'))
            {
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase))
                {
                    List<string> forcesgmts = Regex.Replace(force, "::", "&").Split(':').ToList();

                    foreach (string pin in forcesgmts[1].Split(','))
                    {
                        if (IsAcSpecPin(pin))
                        {
                            acSettings += $"{forcesgmts[0]}:{pin}:{forcesgmts[2]};";
                        }
                    }
                }
            }
            return acSettings.Trim(';');
        }

        public string GetDcSelector()
        {
            string dcAcSelector = "";
            foreach (string force in ForceCondition.Split(';'))
            {
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcSelector, RegexOptions.IgnoreCase))
                {
                    dcAcSelector += force + ";";
                }
            }
            return dcAcSelector.Trim(';');
        }

        public string GetRtosIdsTimeSet()
        {
            foreach (string force in ForceCondition.Split(';'))
            {
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegTimeSet, RegexOptions.IgnoreCase))
                {
                    return force.Split(':')[1];
                }
            }
            return "";
        }

        public string GetMcgSetting()
        {
            string mcgSettings = "";
            foreach (string force in ForceCondition.Split(';'))
            {
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase))
                {
                    List<string> forcesgmts = Regex.Replace(force, "::", "&").Split(':').ToList();

                    foreach (string pin in forcesgmts[1].Split(','))
                    {
                        if (!IsAcSpecPin(pin))
                        {
                            mcgSettings += $"{forcesgmts[0]}:{pin}:{forcesgmts[2]};";
                        }
                    }

                }
            }
            return mcgSettings.Trim(';');
        }

        public string GetPrePatForceCondition()
        {
            // Remove Dc setting, Ac setting, and Mcg setting
            List<string> forcelst = ForceCondition.Split(';').ToList();
            forcelst.RemoveAll(string.IsNullOrEmpty);
            foreach (string force in forcelst.ToArray())
            {
                if (Regex.IsMatch(force, RegLevelSetting, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase)
                    || Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcCategory, RegexOptions.IgnoreCase)
                    || Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcCategory, RegexOptions.IgnoreCase)
                    || Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSelector, RegexOptions.IgnoreCase)
                    || Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcSelector, RegexOptions.IgnoreCase)
                    || Regex.IsMatch(Regex.Replace(force, "::", "&"), RegTimeSet, RegexOptions.IgnoreCase)
                    )
                {
                    forcelst.Remove(force);
                }
            }
            return string.Join(";", forcelst);
        }

        private string CorrectTypo(string inputStr)
        {
            var result = new List<string>();
            foreach (string str in inputStr.Split(';'))
            {
                if (Regex.IsMatch(str.Split(':')[0], "relay_*[(on)|(off)]", RegexOptions.IgnoreCase))
                {
                    List<string> arr = str.Split(':').ToList();
                    if (arr.Count < 2)
                    {
                        result.Add(str);
                    }
                    else
                    {
                        result.Add($"{arr[1]}:{arr[0]}");
                    }
                }
                else
                {
                    result.Add(str);
                }
            }
            return string.Join(";", result);
        }
    }
}
