using System;
using System.Collections.Generic;
using System.Linq;

namespace EfuseCheckCmdLib.IgxlLogLib.Base
{
    public class TouchDown
    {
        public string DeviceInfo;
        public string EnableWords;
        public long RegionStartPtr;
        public long RegionStartLine;
        public long RegionLines;
        public long SegmentSize;
        public long ExecuteTestPtr;
        public long ExecuteTestSize;
        public long SortBinPtr;
        public long SortBinSize;
        public long DieXyPtr;
        public long DieXySize;

        //To fullfill OrangeXL, ConsumerModeOfEachDevice
        public List<int> CurrDeviceNum;
        public List<int> CurrActiveSiteNum;
        //public Dictionary<string, DataLogFormatHeader> CurrentTnFormatFunc;
        //public Dictionary<string, DataLogFormatHeader> CurrentTnFormatMeas;

        private Device[]? _currDevices;
        // each site die infor
        public List<DieInfo> DieTestSum;

        public bool FirstTd { get; set; }

        public Device[] GetDeivices //最後Bin Summary的整理
        {

            get
            {
                if (_currDevices == null)
                {
                    //每個Device一個
                    _currDevices = new Device[CurrActiveSiteNum.Max() + 1];

                    foreach (DieInfo die in DieTestSum)
                    {
                        _currDevices[die.Site] = new Device(die.Site, die.DeviceNum);
                        _currDevices[Convert.ToInt16(die.Site)].FailedTest = Convert.ToInt32(die.ExecFailTests);
                        _currDevices[Convert.ToInt16(die.Site)].ExecutedTest = Convert.ToInt32(die.ExecTests);
                        _currDevices[Convert.ToInt16(die.Site)].Sort = Convert.ToInt32(die.Sort);
                        _currDevices[Convert.ToInt16(die.Site)].Bin = Convert.ToInt32(die.Bin);
                        _currDevices[Convert.ToInt16(die.Site)].LotId = die.LotId;
                        _currDevices[Convert.ToInt16(die.Site)].WfrId = die.WfrId;
                        _currDevices[Convert.ToInt16(die.Site)].EcidCrc = die.EcidCrc;
                        _currDevices[Convert.ToInt16(die.Site)].X = die.XCoord != "N/A" ? Convert.ToInt16(die.XCoord) : -999;
                        _currDevices[Convert.ToInt16(die.Site)].Y = die.YCoord != "N/A" ? Convert.ToInt16(die.YCoord) : -999;
                        _currDevices[Convert.ToInt16(die.Site)].MuliCoreResilt = die.MultiCoreResult;
                    }
                }

                return _currDevices;
            }
        }

        public TouchDown()
        {
            DeviceInfo = "";
            EnableWords = "";
            RegionStartPtr = SegmentSize = ExecuteTestPtr = ExecuteTestSize = 0;
            SortBinPtr = SortBinSize = DieXySize = 0;
            CurrDeviceNum = [];
            CurrActiveSiteNum = [];
            DieTestSum = [];
            FirstTd = false;
        }
    }

}
