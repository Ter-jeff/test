namespace TestPlanLib.BinNumberLegacy
{
    public class BinNumDefRow
    {
        public string Description { set; get; } = "";
        public int SoftBinStart { set; get; } = 0;
        public int SoftBinEnd { set; get; } = 0;
        public string SoftBinState { set; get; } = "Fail";
        public int CurrentSoftBin { set; get; } = 0;
        public string HardBin { set; get; } = "0";
        public string HardIpHvBin { set; get; } = "";
        public string HardIphlvBin { set; get; } = "";
        public string HardIpLvBin { set; get; } = "";
        public string HardIpNvBin { set; get; } = "";
        public SoftBinRangeData CurrentBinLib { set; get; } = new SoftBinRangeData();

        public bool IsExceed { set; get; } = false;
    }
}
