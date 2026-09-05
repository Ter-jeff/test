using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Utility.HardIP;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;

public class CpuFlagAssigner : IFunctionAssigner
{
    public void Assign(Function function, FunctionAssignContext ctx)
    {
        HardIpInfo infoCpuFlag = HardIpService.GetHardIpInfo(ctx.Pattern.Pattern.GetLastPayload());
        string cpuFlag = SearchInfo.GetCpuflag(infoCpuFlag, ctx.Pattern);
        function.SetParamValue("patternCPUAFlag", cpuFlag);
    }
}
