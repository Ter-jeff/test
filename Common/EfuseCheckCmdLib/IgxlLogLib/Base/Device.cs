namespace EfuseCheckCmdLib.IgxlLogLib.Base
{
    public class Device
    {
        //可當作檢索條件 不能重複

        public int DeviceNo;
        public int Site;
        public string SiteDevice { get { return $"Site{Site}:Device{DeviceNo}:X{X}Y{Y}"; } } //只有Site+DeviceNumber都一樣才算重複
        //可當作檢索條件

        public int X = -999;
        //可當作檢索條件
        public int Y = -999;
        public string DieXy { get { return $"{X},{Y}"; } }

        //從結果得到的資訊
        public int Bin = 0;
        public int Sort = 0;
        //<-- 目前發現有的Log沒有
        public int ExecutedTest = -1;
        //<-- 目前發現有的Log沒有
        public int FailedTest = -1;
        public string LotId = "NA";
        public string WfrId = "NA";
        public string EcidCrc = "NA";

        public MutliCore MuliCoreResilt = new();
        //要儲存Test Instance/Number的Sequence!!
        //public List<string> ListSeqTestInstance = new List<string>(); //照順序收集到的Test Instance, 關鍵Key

        public Device() { }

        public Device(int siteNum, int deviceNum)
        {
            Site = siteNum;
            DeviceNo = deviceNum;
        }

    }
}
