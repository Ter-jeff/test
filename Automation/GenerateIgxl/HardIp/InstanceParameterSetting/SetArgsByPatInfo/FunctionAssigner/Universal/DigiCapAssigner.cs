using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class DigiCapAssigner : IFunctionAssigner
{
    public static readonly List<string> ReservedMainProgramCustomString =
    [
        "ConvertDAC",
        "ConvertDACParameter",
        "SignedBinary",
        "RAK",
        "BypassAutorange",
        "DDREmulate",
        "UnSignedGray",
        "SignedGray",
        "TwosComplement",
        "SweepSrc",
        "SweepSrcOneToOne",
        "SweepVoltage",
        "DcvsVoltageClamp",
        "DigCapMsbFirst",
        "DigSrcMsbFirst",
        "RunByModule"
    ];

    public void Assign(Function function, FunctionAssignContext ctx)
    {
        //DigCap_Pin: MeasC pin in Test Plan, Like "MeasC Pin = Pout" ===> Pout
        //function.ArgList[17] = cPin;
        string cPin = SearchInfo.GetMeasCPins(ctx.Pattern, ctx.HardIpInfo);
        function.SetParamValue("digCapPin", cPin);
        //DigCap_DataWidth:  Get from "Cap Bit Str" in patInfo file. Like "wdr14_10+wdr23_10" ===> 10
        //function.ArgList[18] = Regex.Match(info.CapBitStr, @"^wdr\d+_(?<num>(\d+)).*").Groups["num"].ToString();
        function.SetParamValue("digCapDataWidth", SearchInfo.GetDigDataWidth(ctx.HardIpInfo.CapBitStr));
        //DigCap_Sample_Size: Get from "Cap Bit" in patInfo file
        //function.ArgList[19] = info.CapBit.ToString();
        function.SetParamValue("digCapSampleSize", ctx.HardIpInfo.CapBit.ToString("G15"));

        //CUS_Str_DigCapData
        //HardIpInfo hardIpInfo = HardIpService.GetHardIpInfo(ctx.Pattern);
        function.SetParamValue("digCapDataCustomString", SearchInfo.GetCusStrDigCapData(ctx.Pattern, ctx.HardIpInfo, isCsUsing: true));

        var mainProgramCustomStringArg = new List<string>();
        //CUS_Str_MainProgram
        if (ctx.Pattern.Pattern.IsMultiple())
        {
            var dspList = new List<string>();
            foreach (string patternName in ctx.Pattern.Pattern.PatternSetList.SelectMany(p => p).ToList())
            {
                HardIpPattern pat = ctx.Pattern.BurstPatterns.FirstOrDefault(x => x.Pattern.RealPatternName.Equals(patternName));
                if (pat == null)
                {
                    dspList.Add("");
                }
                else
                {
                    dspList.Add(string.Join(";", GetSweepVoltageByVoltage(pat.DspFunction, ctx.Voltage)));
                }
            }
            mainProgramCustomStringArg.Add(string.Join(";", dspList.Where(x => !string.IsNullOrEmpty(x.Trim()))));
        }
        else
        {
            mainProgramCustomStringArg.Add(string.Join(";", GetSweepVoltageByVoltage(ctx.Pattern.DspFunction.Distinct().ToList(), ctx.Voltage)));
        }

        foreach (string reservedKey in ReservedMainProgramCustomString)
        {
            if (!ctx.Pattern.MiscInfoDict.ContainsKey(reservedKey))
            {
                continue;
            }
            if (ctx.Pattern.MiscInfoDict[reservedKey] == "")
            {
                mainProgramCustomStringArg.Add(reservedKey);
                continue;
            }
            foreach (string _ in ctx.Pattern.MiscInfoDict[reservedKey].Split(';'))
            {
                mainProgramCustomStringArg.Add($"{reservedKey}:{ctx.Pattern.MiscInfoDict[reservedKey]}");
            }
        }

        function.SetParamValue("mainProgramCustomString", string.Join(";", mainProgramCustomStringArg.Where(x => !string.IsNullOrEmpty(x))));

        if (ctx.IsAddSweep && ctx.Pattern.ConditionIndex == 0)
        {
            ctx.Pattern.DspFunction.RemoveAt(ctx.Pattern.DspFunction.Count - 1);
        }
    }

    internal static List<string> GetSweepVoltageByVoltage(List<string> strList, string voltage)
    {
        var result = new List<string>();
        foreach (string str in strList)
        {
            var strInside = new List<string>();
            foreach (string subStr in str.Split(';'))
            {
                string newStr = subStr;
                if (voltage.ToUpper() == "NV" && (subStr.StartsWith("LV@sweepvoltage", StringComparison.OrdinalIgnoreCase) || subStr.StartsWith("HV@sweepvoltage", StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (voltage.ToUpper() == "LV" && (subStr.StartsWith("NV@sweepvoltage", StringComparison.OrdinalIgnoreCase) || subStr.StartsWith("HV@sweepvoltage", StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (voltage.ToUpper() == "HV" && (subStr.StartsWith("NV@sweepvoltage", StringComparison.OrdinalIgnoreCase) || subStr.StartsWith("LV@sweepvoltage", StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(voltage) && subStr.StartsWith($"{voltage}@sweepvoltage", StringComparison.OrdinalIgnoreCase))
                {
                    newStr = subStr.Replace($"{voltage}@sweepvoltage", "sweepvoltage");
                }
                else if (string.IsNullOrEmpty(voltage) && subStr.StartsWith("NV@sweepvoltage", StringComparison.OrdinalIgnoreCase))
                {
                    newStr = subStr.Replace("NV@sweepvoltage", "sweepvoltage");
                }
                else if (string.IsNullOrEmpty(voltage) && subStr.StartsWith("LV@sweepvoltage", StringComparison.OrdinalIgnoreCase))
                {
                    newStr = subStr.Replace("LV@sweepvoltage", "sweepvoltage");
                }
                else if (string.IsNullOrEmpty(voltage) && subStr.StartsWith("HV@sweepvoltage", StringComparison.OrdinalIgnoreCase))
                {
                    newStr = subStr.Replace("HV@sweepvoltage", "sweepvoltage");
                }

                strInside.Add(newStr);
            }
            result.Add(string.Join(";", strInside));
        }
        return result;
    }
}
