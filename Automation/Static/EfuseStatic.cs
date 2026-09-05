namespace Automation.Static
{
    public static class EfuseStatic
    {
        public static string HighLightJobs { get; set; } = string.Empty;
        public static int ShowType { get; set; } = 0;
        public static bool IdsRoundMethod { get; set; } = false;
        public static string OutputPath { get; set; } = string.Empty;
        public static bool IsCmd
        {
            get { return !string.IsNullOrEmpty(OutputPath); }
        }
        public static EfuseCheckResultType Result;

        internal static void Clear()
        {
            HighLightJobs = string.Empty;
            ShowType = 0;
            IdsRoundMethod = false;
            OutputPath = string.Empty;
            Result = EfuseCheckResultType.Pass;
        }
    }

    public enum EfuseCheckResultType
    {
        Pass, Fail, Exception
    }
}
