using Automation.Const;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class SpecialCalcAssigner : IFunctionAssigner
{
    public void Assign(Function function, FunctionAssignContext ctx)
    {
        if (ctx.Pattern.SpecialMeasType.Equals(MeasType.MeasVdiff))
        {
            function.SetParamValue("specialCalcSetting", "4");
        }
        else if (ctx.Pattern.SpecialMeasType.Equals(MeasType.MeasVocm))
        {
            function.SetParamValue("specialCalcSetting", "9");
        }
    }
}
