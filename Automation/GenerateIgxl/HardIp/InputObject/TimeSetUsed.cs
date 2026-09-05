namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class TimeSetUsed
    {
        public string McgSetting { get; set; }
        public string TimeSet { get; set; }
        public string TimeSetMcg { get; set; }

        public TimeSetUsed()
        {
        }

        public TimeSetUsed(TimeSetUsed other)
        {
            if (other == null)
            {
                return;
            }

            McgSetting = other.McgSetting;
            TimeSet = other.TimeSet;
            TimeSetMcg = other.TimeSetMcg;
        }
    }
}
