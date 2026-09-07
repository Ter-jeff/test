using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using TestPlanLib.Basic;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetIdsMathFuncValue : SetValueBase
    {
        private List<string> _miscInfoList = new List<string>();
        public SetIdsMathFuncValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
            ReservedMiscInfoKeys = new List<string>
            {
                "Current1_dic",
                "Current2_dic",
                "power_rail",
                "Calc_Pin",
                "MathFunction",
                "MathInverse",
                "FusedStage"
            };
        }
        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            _miscInfoList = pattern.MiscInfo.Split(';').ToList();
            if (pattern.MiscInfo.IndexOf("Func:IDS_Delta_calc", StringComparison.OrdinalIgnoreCase) != -1)
            {
                SetForDelta(function);
            }
            else if (pattern.MiscInfo.IndexOf("Func:IDS_main_current_MathFunc", StringComparison.OrdinalIgnoreCase) != -1)
            {
                SetForMath(function);
            }
        }
        public void SetForDelta(Function function)
        {
            string current1Dic = _miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("Current1_dic", StringComparison.OrdinalIgnoreCase)) ?? "";
            string current2Dic = _miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("Current2_dic", StringComparison.OrdinalIgnoreCase)) ?? "";
            string powerRail = _miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("power_rail", StringComparison.OrdinalIgnoreCase)) ?? "";

            if (powerRail.Split(':').Length > 1)
            {
                powerRail = powerRail.Split(':')[1].Trim();
            }
            if (current1Dic.Split(':').Length > 1)
            {
                current1Dic = current1Dic.Split(':')[1].Trim();
                if (!string.IsNullOrEmpty(current1Dic) && !string.IsNullOrEmpty(powerRail))
                {
                    function.SetParamValue("firstCalculatedData", $"{current1Dic}:{powerRail}");
                }
            }
            if (current2Dic.Split(':').Length > 1)
            {
                current2Dic = current2Dic.Split(':')[1].Trim();
                if (!string.IsNullOrEmpty(current2Dic) && !string.IsNullOrEmpty(powerRail))
                {
                    function.SetParamValue("secondCalculatedData", $"{current2Dic}:{powerRail}");
                }
            }
            function.SetParamValue("operationSymbol", "-");
        }
        public void SetForMath(Function function)
        {
            string calcPin = _miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("Calc_Pin", StringComparison.OrdinalIgnoreCase)) ?? "";
            string mathFunction = _miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("MathFunction", StringComparison.OrdinalIgnoreCase)) ?? "";
            string mathInverse = _miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("MathInverse", StringComparison.OrdinalIgnoreCase)) ?? "";
            string fusedStage = _miscInfoList.FirstOrDefault(x => x.Split(':')[0].Trim().Equals("FusedStage", StringComparison.OrdinalIgnoreCase)) ?? "";
            string fuseStrFromCalcIdsSum = GetFuseStrFromCalcIdsSum().ToUpper();
            KeyValuePair<string, string> calcIdsSum = fuseStrFromCalcIdsSum.Split('#').Length > 1 ?
                new KeyValuePair<string, string>(fuseStrFromCalcIdsSum.Split('#')[0], fuseStrFromCalcIdsSum.Split('#')[1]) : new KeyValuePair<string, string>("", "");

            string measStr = "";
            string fuseStr = "";
            bool inverse = false;

            if (mathFunction.Split(':').Length > 1)
            {
                mathFunction = mathFunction.Split(':')[1].Trim();
                function.SetParamValue("operationSymbol", mathFunction);
            }
            if (calcPin.Split(':').Length > 1)
            {
                calcPin = calcPin.Split(':')[1].Trim();
                measStr = $"ON:{calcPin}";
            }
            if (mathInverse.Split(':').Length > 1)
            {
                inverse = mathInverse.Split(':')[1].Trim().ToLower() == "true";
            }
            if (fusedStage.Split(':').Length > 1)
            {
                fusedStage = fusedStage.Split(':')[1].Trim().ToUpper();
                if (!string.IsNullOrEmpty(fusedStage))
                {
                    var fusePins = new List<string>();
                    var fuseMappingList = TestPlanStatic.IdsMappingSheet.Rows.Where(x => x.Stage.Trim().ToUpper() == fusedStage).ToList();
                    foreach (string pin in calcPin.Split(','))
                    {
                        if (calcIdsSum.Key != "" && calcIdsSum.Key == pin.ToUpper())
                        {
                            fusePins.Add(calcIdsSum.Value);
                        }
                        else
                        {
                            IdsMappingRow fuseMapping = fuseMappingList.FirstOrDefault(x => x.Pinname.Trim().ToUpper() == pin.Trim().ToUpper());
                            if (fuseMapping != null)
                            {
                                fusePins.Add($"{fuseMapping.Efusefieldname.Trim().ToUpper()}");
                            }
                        }
                    }
                    fuseStr = $"{fusedStage}:{string.Join(",", fusePins)}";
                }
            }
            if (inverse)
            {
                function.SetParamValue("firstCalculatedData", fuseStr);
                function.SetParamValue("secondCalculatedData", measStr);
            }
            else
            {
                function.SetParamValue("firstCalculatedData", measStr);
                function.SetParamValue("secondCalculatedData", fuseStr);
            }
        }

        private string GetFuseStrFromCalcIdsSum()
        {
            string result = "";
            var fuseStrRegex = new Regex(@"Calc_IDS_Sum\(EFUSE@(?<FuseSumStr>.*)\)", RegexOptions.IgnoreCase);

            foreach (string miscInfo in _miscInfoList)
            {
                Match m = fuseStrRegex.Match(miscInfo);
                if (m.Success)
                {
                    result = m.Groups["FuseSumStr"].Value;
                    break;
                }
            }
            return result;
        }
    }
}
