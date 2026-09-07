using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Utility.HardIP;

using CommonLib.Extension;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetVirValue : SetValueBase
    {
        public SetVirValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            string cPin = SearchInfo.GetMeasCPins(pattern, HardIpService.GetHardIpInfo(pattern));
            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);

            //patset
            function.SetParamValue("patset", pattern.Pattern.GetInstancePatternName());
            //Cpu_flag_A
            //Meas_VIR_IO_Universal_func_GPIO_TTR
            if (function.FunctionName.ContainsIgnoreCase("gpio_ttr"))
            {
                function.SetParamValue("CPUA_Flag_In_Pat", "FALSE");
            }
            else
            {
                function.SetParamValue("CPUA_Flag_In_Pat", SearchInfo.GetCpuflag(info, pattern));
            }

            //TestLimitPerPin
            function.SetParamValue("TestLimitPerPin_VIR", string.Join("", SearchInfo.GetTestLimitPerMeasType(pattern).Values.ToList()));

            #region RegisterAssignment
            //DigSrc_Assignment: Use "Register Assignment" value in test plan
            function.SetParamValue("DigSrc_Assignment", pattern.RegisterAssignment);
            //DigSrc_Equation: From patInfo file "Send Bit Name"
            function.SetParamValue("DigSrc_Equation", pattern.DigSrcEquation);
            //DigSrc_Sample_Size: Get from "Send Bit" in patInfo file, Like Send Bit: 160  ===> 160
            function.SetParamValue("DigSrc_Sample_Size", info.SendBit);
            //DigSrc_DataWidth: Get from "Send Bit Str" in patInfo file. Like wdr0_16+wdr1_16+wdr2_16 ===> 16
            function.SetParamValue("DigSrc_DataWidth", SearchInfo.GetDigDataWidth(info.SendBitStr, "0"));
            //DigSrc_Pin
            function.SetParamValue("DigSrc_Pin", SearchInfo.GetSrcPin(info));
            #endregion

            #region MeasC
            //DigCap_Pin: MeasC pin in Test Plan, Like "MeasC Pin = Pout" ===> Pout
            function.SetParamValue("DigCap_Pin", cPin);
            //DigCap_DataWidth:  Get from "Cap Bit Str" in patInfo file. Like "wdr14_10+wdr23_10" ===> 10
            function.SetParamValue("DigCap_DataWidth", SearchInfo.GetDigDataWidth(info.CapBitStr));
            //DigCap_Sample_Size: Get from "Cap Bit" in patInfo file
            function.SetParamValue("DigCap_Sample_Size", info.CapBit.ToString("G15"));
            //CUS_Str_DigCapData
            HardIpInfo hardIpInfo = HardIpService.GetHardIpInfo(pattern);
            function.SetParamValue("CUS_Str_DigCapData", SearchInfo.GetCusStrDigCapData(pattern, hardIpInfo));
            #endregion

            function.SetParamValue("Measure_Pin_PPMU", SearchInfo.GetPpmuPin(pattern, info));
            string forceV = SearchInfo.GetForceV(pattern, info, voltage);
            function.SetParamValue("ForceV", forceV);
            function.SetParamValue("ForceI", SearchInfo.GetForceI(pattern, info, voltage));
            function.SetParamValue("MeasureI_Range", SearchInfo.GetIRange(pattern, info, voltage));

            //Interpose_PrePat and Iterpose_PostTest
            function.SetParamValue("Interpose_PrePat", SearchInfo.GetPrePat(pattern, voltage));//Use "Interpose_PrePat" to instead of "CharInputString"
            function.SetParamValue("Interpose_PreMeas", SearchInfo.GetPreMeas(pattern, HardIpInputData, "", voltage));
            function.SetParamValue("Interpose_PostMeas", SearchInfo.GetPostMeas(pattern));
            function.SetParamValue("Interpose_PostTest", pattern.InterposePostTest);

            //Measure Sequence
            string measSeq = SearchInfo.GetMeasSequence(pattern).Replace("R1", "R").Replace("VDIFF", "V").Replace("IDIFF", "I").Replace("R2", "Z");

            function.SetParamValue("TestSequence", measSeq);

            //SpecialCalcValSetting

            if (pattern.SpecialMeasType.Equals(MeasType.MeasVdiff))
            {
                function.SetParamValue("SpecialCalcValSetting", "4");
            }

            if (pattern.SpecialMeasType.Equals(MeasType.MeasVocm))
            {
                function.SetParamValue("SpecialCalcValSetting", "9");
            }

            //Calc_Eqn and storeName
            function.SetParamValue("Calc_Eqn", DataConvertor.ConvertValueSpec(pattern.CalcEqn));
            string storeNameOri = "";
            function.SetParamValue("Meas_StoreName", DataConvertor.SortCpFtPin(SearchInfo.GetStoreName(pattern, HardIpInputData, ref storeNameOri, measSeq)));

            //SweepCode
            function.SetParamValue("DigSrc_FlowForLoopIntegerName", HardIpService.GetFlowForLoopIntegerName(pattern, HardIpInputData));

            //Set MeasIGrpPinCnt MeasVGrpPinCnt to 1 for
            function.SetParamValue("MeasIGrpPinCnt", "1");
            function.SetParamValue("MeasVGrpPinCnt", "1");
        }

    }
}
