using EfuseCheckCmdLib.EFuse.EFuseApp;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    public class XBitDefRow
    {
        //from output
        public int RowNo;
        public string RealStage = "";
        public int JobStage;
        public string Name = "";
        public string Resolution = "";
        public bool IsDefault;
        public string DefaultReal = "";
        public string Msb = "";
        public string Lsb = "";
        public double LowLimit;
        public double HighLimit;
        public double RealValue;
        public string DefaultValue = "";
        public string Block = "";
        public string Algorithm = "";
        public string HiPparName = "";
        public string HiPparEq = "";
        public double HipHighLimit;
        public double HipLowLimit;
        public EnumValueType Type;
    }
}
