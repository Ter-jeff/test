using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class InterposeAssigner : IFunctionAssigner
{
    public void Assign(Function function, FunctionAssignContext ctx)
    {
        //Interpose_PrePat,Premeas,PostTest
        string prePat = SearchInfo.GetPrePat(ctx.Pattern, ctx.Voltage);
        function.SetParamValue("interposePrePat", prePat);
        string preMeas = SearchInfo.GetPreMeas(ctx.Pattern, ctx.InputData, ctx.TestSequence, ctx.Voltage);
        function.SetParamValue("interposePreMeas", preMeas);
        string postMeas = SearchInfo.GetPostMeas(ctx.Pattern);
        function.SetParamValue("interposePostMeas", postMeas);
        function.SetParamValue("interposePostTest", ctx.Pattern.InterposePostTest);
    }
}
