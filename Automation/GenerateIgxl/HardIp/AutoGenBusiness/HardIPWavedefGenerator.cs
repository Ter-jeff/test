using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;

using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class HardIpWavedefGenerator
    {
        public WaveDefinitionSheet Generate(Dictionary<string, HardIpSheet> planDic)
        {
            WaveDefinitionSheet sheet = new WaveDefinitionSheet("Wave Def");
            return sheet;
        }
    }
}
