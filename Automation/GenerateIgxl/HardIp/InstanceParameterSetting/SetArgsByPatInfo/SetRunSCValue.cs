using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetRunScValue : SetValueBase
    {
        private Dictionary<string, string> _miscInfoDict = new Dictionary<string, string>();

        public SetRunScValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
            ReservedMiscInfoKeys = new List<string>
            {
                "cmdTimeoutMs"
            };
        }

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            TimeoutProcess(pattern, function);
        }

        private void TimeoutProcess(HardIpPattern pattern, Function function)
        {
            double totalTimeout = 0.0;
            _miscInfoDict = pattern.MiscInfoDict;
            foreach (KeyValuePair<string, string> para in _miscInfoDict)
            {
                if (para.Key.ToUpper().EndsWith("TIMEOUT"))
                {
                    double.TryParse(para.Value, out double timeout);
                    if (timeout != -1.0)
                    {
                        totalTimeout += timeout;
                    }
                }
            }
            function.SetParamValue("cmdTimeoutMs", totalTimeout.ToString());
        }
    }
}
