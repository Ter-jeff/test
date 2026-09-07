using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetFunctionalValue : SetValueBase
    {
        private readonly Regex _selSramRegex = new Regex(@"_D?SRA?M\w*DSSC", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public SetFunctionalValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            if (function.Type == ".NET")
            {
                CsProcess(pattern, function, voltage);
            }
            else
            {
                VbtProcess(pattern, function, voltage);
            }
        }
        private void CsProcess(HardIpPattern pattern, Function function, string voltage)
        {
            function.CheckParam = false;

            #region Set value for Functional_T_Updated
            //Patterns
            function.SetParamValue("patterns", pattern.Pattern.IsMultiple() ?
                                                string.Join(",", pattern.Pattern.InstancePayloadName) :
                                                pattern.Pattern.GetInstancePatternName());

            #region Default value
            //RelayMode
            function.SetParamValue("relayMode", "1");
            function.SetParamValue("patternTimeout", "30");
            #endregion

            //DigSrc
            string digSrcPin = "";
            string digSrcEquation = "";
            string digSrcAssignment = "";
            GetDigSrc(pattern, ref digSrcPin, ref digSrcEquation, ref digSrcAssignment);
            function.SetParamValue("digSrcPin", digSrcPin);
            function.SetParamValue("digSrcEquation", digSrcEquation);
            function.SetParamValue("digSrcAssignment", digSrcAssignment);

            string interposePrePat = SearchInfo.GetPrePat(pattern, voltage);
            if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                SetInterPoseFunc(pattern, function, voltage);
            }
            else
            {
                function.SetParamValue("interposePrepat", interposePrePat);
                function.SetParamValue("Interpose_PostMeas", SearchInfo.GetPostMeas(pattern));
            }
            function.SetParamValue("resultMode", pattern.Pattern.IsMultiple() ? "1" : "0");

            #endregion
        }
        private void VbtProcess(HardIpPattern pattern, Function function, string voltage)
        {
            function.CheckParam = false;

            #region Set value for Functional_T_Updated
            //Patterns
            function.SetParamValue("Patterns", pattern.Pattern.IsMultiple() ?
                string.Join(",", pattern.Pattern.InstancePayloadName) :
                pattern.Pattern.GetInstancePatternName());
            #region Default value
            //RelayMode
            function.SetParamValue("RelayMode", "1");
            function.SetParamValue("PatternTimeout", "30");
            function.SetParamValue("DoApplyLevelsTiming", "TRUE");
            #endregion

            if (!string.IsNullOrEmpty(pattern.BlockType) && pattern.Pattern.PatternSetList.Last().Exists(x => Regex.IsMatch(x, @"_D*SRA*M\w*DSSC", RegexOptions.IgnoreCase)))
            {
                function.SetParamValue("DigSource", Function.TestAutoSwitchJtagTdi);
            }

            string interposePrePat = SearchInfo.GetPrePat(pattern, voltage);
            if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                SetInterPoseFunc(pattern, function, voltage);
            }
            else
            {
                function.SetParamValue("Interpose_PrePat", interposePrePat);
                function.SetParamValue("Interpose_PostMeas", SearchInfo.GetPostMeas(pattern));
            }
            function.SetParamValue("CharInputString", interposePrePat);
            function.SetParamValue("Resultmode", pattern.Pattern.IsMultiple() ? "1" : "0");

            #endregion
        }
        internal void GetDigSrc(HardIpPattern pattern, ref string digSrcPin, ref string digSrcEquation, ref string digSrcAssignment)
        {
            var eqnList = new List<string>();
            var assignments = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var burstPatterns = pattern.Pattern.PatternSetList.SelectMany(p => p).ToList();

            foreach (string burstPattern in burstPatterns)
            {
                if (_selSramRegex.IsMatch(burstPattern))
                {
                    eqnList.Add("C");
                    assignments.Add("C=Selsram()");
                }
                else
                {
                    eqnList.Add("");
                }
            }
            if (eqnList.Any(x => !string.IsNullOrEmpty(x)))
            {
                digSrcPin = "JTAG_TDI";
                digSrcEquation = string.Join("|", eqnList);
            }
            digSrcAssignment = string.Join(";", assignments);
        }
    }
}
