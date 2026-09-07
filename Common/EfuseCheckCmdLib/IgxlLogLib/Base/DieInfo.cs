namespace EfuseCheckCmdLib.IgxlLogLib.Base
{
    public class DieInfo
    {
        public int DeviceNum;
        public int Site;
        public string ExecFailTests = "";
        public string ExecTests = "";
        public string Sort = "";
        public string Bin = "";
        public string XCoord = "";
        public string YCoord = "";
        public string LotId = "";
        public string WfrId = "";
        public string EcidCrc = "";
        public MutliCore MultiCoreResult = new();
    }
}
