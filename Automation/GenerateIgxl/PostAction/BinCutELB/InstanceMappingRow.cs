using System.Collections.Generic;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.PostAction.BinCutELB
{
    public class InstanceMappingRow
    {
        public List<string> BinCutInstanceNames = new List<string>();
        public string HardipInstanceName;
        public bool IsFound;
        public InstanceRow BinCutRow;
    }
}
