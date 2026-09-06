using System.Collections.Generic;

namespace EfuseCheckCmdLib.IgxlLogLib.Base
{
    public class DeviceData(int siteNum, int deviceNum)
    {

        public int SiteNum = siteNum;
        public int DeviceNum = deviceNum;
        public List<ILogRow> DataLowRows = [];
    }
}
