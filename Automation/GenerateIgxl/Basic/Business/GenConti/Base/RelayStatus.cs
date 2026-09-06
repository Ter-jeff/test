using System.Collections.Generic;
using System.Linq;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base
{
    public class RelayStatus
    {
        public List<string> OpenRelayList { set; get; } = new List<string>();
        public List<string> OffRelayList { set; get; } = new List<string>();

        public bool IsEqualStatus(RelayStatus targetStatus)
        {
            if (!OpenRelayList.Any() && !targetStatus.OpenRelayList.Any())
            {
                return true;
            }
            if (OpenRelayList.All(targetStatus.OpenRelayList.Contains) &&
                targetStatus.OpenRelayList.All(OpenRelayList.Contains))
            {
                return true;
            }
            return false;
        }
    }
}
