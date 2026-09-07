using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Utility.HardIP;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetDefaultValue : SetValueBase
    {
        public SetDefaultValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {

            #region If the function contains recognised paramters, try to set value for them

            HardIpService.GetHardIpInfo(pattern.Pattern.GetLastPayload());
            function.CheckParam = false;

            SetValueBase setVbt = new SetVifValue(HardIpInputData, HardIpSheet);
            setVbt.SetArgsListValue(pattern, ref function, voltage);

            #endregion
        }
    }
}
