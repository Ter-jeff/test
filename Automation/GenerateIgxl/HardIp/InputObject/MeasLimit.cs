namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class MeasLimit
    {
        public string JobName { get; set; }
        public string LoLimit { get; set; }
        public string HiLimit { get; set; }
        public int LoHeaderIndex { get; set; }
        public int HiHeaderIndex { get; set; }
        public MeasLimit(string jobName)
        {
            JobName = jobName;
            LoLimit = "";
            HiLimit = "";
        }
    }
}
