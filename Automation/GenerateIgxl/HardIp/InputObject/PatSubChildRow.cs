using System.Collections.Generic;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class PatSubChildRow : PatChildRow
    {
        public List<TestPlanRow> TpRows { get; set; } = new List<TestPlanRow>();
    }
}
