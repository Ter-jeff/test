using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner;

public interface IFunctionAssigner
{
    void Assign(Function function, FunctionAssignContext ctx);
}
