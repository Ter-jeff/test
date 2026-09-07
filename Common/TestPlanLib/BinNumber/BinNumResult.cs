namespace TestPlanLib.BinNumber
{
    public class BinNumResult(int softBin, BinNumInfo binNumInfo, bool found)
    {
        public int SoftBin { get; private set; } = softBin;
        public BinNumInfo BinNumInfo { get; private set; } = binNumInfo;
        public bool Found { get; set; } = found;
    }
}
