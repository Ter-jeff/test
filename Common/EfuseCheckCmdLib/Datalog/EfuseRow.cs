namespace EfuseCheckCmdLib.Datalog
{
    public class EfuseRow
    {
        public string TestNames { get; set; } = "";
        public string BankConfig { get; set; } = "";
        public string SubConfig { get; set; } = "";
        public int MsbBit { get; set; }
        public int LsbBit { get; set; }
        public string ProgrammingStage { get; set; } = "";
        public int BitWidth { get; set; }
        public string LowLimit { get; set; } = "";
        public string HighLimit { get; set; } = "";
        public string Resolution { get; set; } = "";
        public string Algorithm { get; set; } = "";
        public string Description { get; set; } = "";
        public string DefaultOrReal { get; set; } = "";
        public string HexValue { get; set; } = "";
        public string DecValue { get; set; } = "";
        public string BinValue { get; set; } = "";
        public int Site { get; set; }

        public EfuseLine Line { get; set; } = null!;
        public string Bits { get; internal set; } = "";
        public string SortedBits { get; internal set; } = "";
        public string Data { get; internal set; } = "";
    }
}
