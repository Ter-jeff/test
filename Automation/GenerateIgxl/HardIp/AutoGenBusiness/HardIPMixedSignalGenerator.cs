using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;

using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class HardIpMixedSignalGenerator
    {
        public MixedSignalSheet Generate(Dictionary<string, HardIpSheet> planDic)
        {
            MixedSignalSheet sheet = new MixedSignalSheet("Mixed Signal Timing");
            return sheet;
        }
    }
}
