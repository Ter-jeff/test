using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class PatSetAssigner : IFunctionAssigner
{
    public void Assign(Function function, FunctionAssignContext ctx)
    {
        string patSetValue = ctx.Pattern.Pattern.IsMultiple()
            ? string.Join(",", ctx.Pattern.Pattern.InstancePayloadName)
            : ctx.Pattern.Pattern.GetInstancePatternName();
        function.SetParamValue("patternSet", patSetValue);
    }
}
