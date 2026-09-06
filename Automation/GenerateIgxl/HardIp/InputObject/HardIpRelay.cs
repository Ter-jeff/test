namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class HardIpRelay
    {
        public string Job { get; }

        public string Name { get; }

        public RelayStatus Status { get; }

        public HardIpRelay(string job, string name, RelayStatus status)
        {
            Job = job;
            Name = name;
            Status = status;
        }

    }

    public enum RelayStatus
    {
        On = 0,
        Off
    }
}
