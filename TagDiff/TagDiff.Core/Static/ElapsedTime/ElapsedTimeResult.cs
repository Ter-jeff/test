namespace TagDiff.Core.Static.ElapsedTime
{
    public class ElapsedTimeResult(string module, string time)
    {
        public string Module { get; set; } = module;
        public string Time { get; set; } = time;
    }
}
