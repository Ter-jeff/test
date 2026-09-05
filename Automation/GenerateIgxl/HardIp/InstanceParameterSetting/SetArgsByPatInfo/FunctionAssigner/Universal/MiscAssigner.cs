using System;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.InputManager.Data;
using Automation.Utility.HardIP;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class MiscAssigner : IFunctionAssigner
{
    public void Assign(Function function, FunctionAssignContext ctx)
    {
        //Calc_Eqn and storeName
        function.SetParamValue("calcEquation", DataConvertor.ConvertValueSpec(ctx.Pattern.CalcEqn, ctx.Voltage));

        if (ctx.Pattern.MiscInfo.Split(';').Any(
            p => p.Split(':')[0].Equals(HardIpConstData.InstSpecialSetup, StringComparison.OrdinalIgnoreCase)))
        {
            string specialsetup = ctx.Pattern.MiscInfo.Split(';').FirstOrDefault(
                p => p.Split(':')[0].Equals(HardIpConstData.InstSpecialSetup, StringComparison.OrdinalIgnoreCase));
            function.SetParamValue(HardIpConstData.InstSpecialSetup, GetInstSpecialSetup(specialsetup, ctx.InputData));
        }

        if (SearchInfo.GetFlagSingleLimit(ctx.Pattern, ctx.Voltage))
        {
            function.SetParamValue("singleLimitFlag", "TRUE");
        }

        //SweepCode
        function.SetParamValue("digSrcFlowForLoopIntegerName", HardIpService.GetFlowForLoopIntegerName(ctx.Pattern, ctx.InputData, isCsUsing: true));
        function.SetParamValue("setupMSBFirstFlag", ctx.HardIpInfo.IsMsbFirst());

        //digSrcDataCustomString for SelSramBlock
        if (ctx.Pattern.MiscInfoDict.TryGetValue("SelSramBlock", out string value))
        {
            function.SetParamValue("digSrcDataCustomString", $"{"SelSramBlock"}:{ctx.Pattern.MiscInfoDict["SelSramBlock"].Split(';')[0]}");
        }
    }

    internal static string GetInstSpecialSetup(string setup, HardIpInputData inputData)
    {
        string result = "";
        if (setup.Contains(':'))
        {
            string value = setup.Split(':')[1].Trim();
            if (int.TryParse(value, out int _))
            {
                result = value;
            }
            else
            {
                if (inputData.ConfigData.InstSpecialSetting.TryGetValue(value, out string value1))
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
