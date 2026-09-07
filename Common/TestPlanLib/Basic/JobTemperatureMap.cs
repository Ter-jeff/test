namespace TestPlanLib.Basic
{
    public class JobTemperatureMap(string jobName, string temperature, string testStage)
    {
        public const string Rt = "RT";
        public const string Ht = "HT";
        public const string Cp = "CP";
        public const string Ft = "FT";

        public string JobName { set; get; } = jobName;
        public string Temperature { set; get; } = temperature;
        public string TestStage { set; get; } = testStage;
    }
}
