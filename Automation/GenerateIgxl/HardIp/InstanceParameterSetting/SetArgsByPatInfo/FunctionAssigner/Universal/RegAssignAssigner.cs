using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Utility.HardIP;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class RegAssignAssigner : IFunctionAssigner
{
    public void Assign(Function function, FunctionAssignContext ctx)
    {

        //DigSrc_Assignment: Use "Register Assignment" value in test plan
        //function.ArgList[25] = pattern.RegisterAssignment;
        function.SetParamValue("digSrcAssignment", ctx.DigSrcAssignmentVal);

        bool isMultiple = ctx.Pattern.Pattern.IsMultiple();
        List<string> equations = [];
        List<string> sendBits = [];
        if (isMultiple)
        {
            List<string> dataWidths = [];
            var patternNames = ctx.Pattern.Pattern.PatternSetList.SelectMany(p => p).ToList();
            foreach (string patternName in patternNames)
            {
                HardIpInfo multiInfo = HardIpService.GetHardIpInfo(patternName);
                equations.Add(multiInfo.SendBitName);
                sendBits.Add(multiInfo.SendBit);
                dataWidths.Add(SearchInfo.GetDigDataWidth(multiInfo.SendBitStr, "0"));
            }

            //DigSrc_Equation: From patInfo file "Send Bit Name"
            ctx.Pattern.DigSrcEquation = equations.Any(x => !string.IsNullOrEmpty(x))
                ? string.Join("|", equations)
                : string.Empty;

            //DigSrc_Sample_Size: Get from "Send Bit" in patInfo file, Like Send Bit: 160  ===> 160
            //function.ArgList[23] = info.SendBit.ToString();
            if (sendBits.Any(p => !string.IsNullOrEmpty(p)))
            {
                ctx.HardIpInfo.SendBit = string.Join("|", sendBits);
            }
            //DigSrc_DataWidth: Get from "Send Bit Str" in patInfo file. Like wdr0_16+wdr1_16+wdr2_16 ===> 16
            //function.ArgList[22] =
            if (dataWidths.Any(p => !string.IsNullOrEmpty(p)))
            {
                function.SetParamValue("digSrcDataWidth", string.Join("|", dataWidths));
            }
        }
        else
        {
            function.SetParamValue("digSrcDataWidth", SearchInfo.GetDigDataWidth(ctx.HardIpInfo.SendBitStr, "0"));
        }

        function.SetParamValue("digSrcEquation", ctx.Pattern.DigSrcEquation);
        function.SetParamValue("digSrcSampleSize", ctx.HardIpInfo.SendBit);

        //DigSrc_Pin
        function.SetParamValue("digSrcPin", SearchInfo.GetSrcPin(ctx.HardIpInfo));
    }
}
