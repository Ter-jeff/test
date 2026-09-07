using Automation.GenerateIgxl.HardIp.InputObject;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public abstract class HardIpPrecheckBase
    {
        protected readonly HardIpSheet HardIpSheet;
        protected readonly HardIpPattern Pattern;

        protected HardIpPrecheckBase(HardIpSheet hardIpSheet, HardIpPattern pattern)
        {
            HardIpSheet = hardIpSheet;
            Pattern = pattern;
        }

        public abstract void Check();
    }
}
