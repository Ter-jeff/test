using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class MeasCAssigner : IFunctionAssigner
{
    public void Assign(Function function, FunctionAssignContext ctx)
    {
        function.SetParamValue("measureStoreName", ctx.MeasStoreName);

        string forceV = SearchInfo.GetForceV(ctx.Pattern, ctx.HardIpInfo, ctx.Voltage, true);
        string forceI = SearchInfo.GetForceI(ctx.Pattern, ctx.HardIpInfo, ctx.Voltage);
        function.SetParamValue("measureForceV", CheckAddSymbol(SearchInfo.CheckInfoByStoreName(forceV, ctx.OriginalStoreName, '|')));
        function.SetParamValue("measureForceI", CheckAddSymbol(SearchInfo.CheckInfoByStoreName(forceI, ctx.OriginalStoreName, '|')));
        //Measure Sequence 
        function.SetParamValue("measureSequence", ctx.TestSequence);
    }

    private static string CheckAddSymbol(string input)
    {
        if (Regex.IsMatch(input, "[:+]", RegexOptions.IgnoreCase))
        {
            List<string> sgmts = Regex.Split(input, "[:+]", RegexOptions.IgnoreCase).ToList();
            if (sgmts.All(p => double.TryParse(p, out double _)))
            {
                return "@" + input;
            }
        }
        return input;
    }
}
