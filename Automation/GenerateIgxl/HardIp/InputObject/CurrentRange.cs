namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class CurrentRange
    {
        public string JobName { get; set; }
        public string Value { get; set; }

        public CurrentRange()
        {
            JobName = "";
            Value = "";
        }

        public CurrentRange(string jobName, string value)
        {
            JobName = jobName;
            Value = value;
        }
    }
}
