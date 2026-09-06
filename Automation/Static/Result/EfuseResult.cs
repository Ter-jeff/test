namespace Automation.Static.Result
{
    public class EfuseResult
    {
        public static bool HasBkmProcess { get; set; }

        public static void Clear()
        {
            HasBkmProcess = false;
        }
    }
}
