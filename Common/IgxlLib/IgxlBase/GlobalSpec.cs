namespace IgxlLib.IgxlBase
{
    public class GlobalSpec : Spec
    {
        public string Job { get; set; } = string.Empty;

        public GlobalSpec(string symbol) : base(symbol)
        {
        }

        public GlobalSpec(string symbol, string value) : base(symbol, value)
        {
        }

        public GlobalSpec(string symbol, string job, string value, string comment) : base(symbol, value, comment)
        {
            Job = job;
        }
    }
}
