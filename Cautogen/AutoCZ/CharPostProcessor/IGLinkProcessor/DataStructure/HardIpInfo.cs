using System.Collections.Generic;

namespace Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure
{
    public class HardIpSeqInfo
    {
        // properties
        public string SeqName { get; set; }
        public string PinList { get; set; }
        public List<string> MeasIPowerPinList = new List<string>();
        public List<string> MeasVPowerPinList = new List<string>();
        public List<string> MeasIioPinList = new List<string>();
        public List<string> MeasVioPinList = new List<string>();
        public List<string> MeasFPinList = new List<string>();
        public List<string> MeasVdiffPinList = new List<string>();
        public List<string> MeasIdiffPinList = new List<string>();
        public List<string> MeasVdiff2PinList = new List<string>();
        public List<string> MeasFdiffPinList = new List<string>();
        public List<string> MeasVocmPinList = new List<string>();
        public List<string> MeasR1PinList = new List<string>();
        public List<string> MeasR2PinList = new List<string>();

        // constructor
        public HardIpSeqInfo()
        {
            SeqName = "";
            PinList = "";
        }
    }


    public class HardIpReference
    {
        // properties
        public string Payload { get; set; }
        public string Vm { get; set; }
        public string TimeSet { get; set; }
        public string Subr { get; set; }
        public int CapBit { get; set; }
        public int SendBit { get; set; }
        public string CapBitStr { get; set; }
        public string SendBitStr { get; set; }
        public string CapBitName { get; set; }
        public string SendBitName { get; set; }
        public string CapPinName { get; set; }
        public string SendPinName { get; set; }
        public string DsscOut { get; set; }
        public string MeasIStr { get; set; }
        public string MeasVStr { get; set; }
        public string MeasFStr { get; set; }
        public string MeasVdiffStr { get; set; }
        public string MeasVdiff2Str { get; set; }
        public string MeasIdiffStr { get; set; }
        public string MeasFdiffStr { get; set; }
        public string MeasVocmStr { get; set; }
        public string MeasR1Str { get; set; }
        public string MeasR2Str { get; set; }
        public string MeasSeqStr { get; set; }
        public string DigSrcAssign { get; set; }
        public string DigSrcDataWidth { get; set; }
        public bool IsIgnoreComment { get; set; }
        public List<string> MeasIPowerPinList = new List<string>();
        public List<string> MeasVPowerPinList = new List<string>();
        public List<string> MeasIioPinList = new List<string>();
        public List<string> MeasVioPinList = new List<string>();
        public List<string> MeasFPinList = new List<string>();
        public List<string> MeasVdiffPinList = new List<string>();
        public List<string> MeasIdiffPinList = new List<string>();
        public List<string> MeasVdiff2PinList = new List<string>();
        public List<string> MeasFdiffPinList = new List<string>();
        public List<string> MeasVocmPinList = new List<string>();
        public List<string> MeasR1PinList = new List<string>();
        public List<string> MeasR2PinList = new List<string>();
        public List<HardIpSeqInfo> SeqInfo = new List<HardIpSeqInfo>();

        // constructor
        public HardIpReference()
        {
            DigSrcDataWidth = "";
            DigSrcAssign = "";
            MeasSeqStr = "";
            MeasR2Str = "";
            MeasR1Str = "";
            MeasVocmStr = "";
            MeasFdiffStr = "";
            MeasIdiffStr = "";
            MeasVdiff2Str = "";
            MeasVdiffStr = "";
            MeasFStr = "";
            MeasVStr = "";
            MeasIStr = "";
            DsscOut = "";
            SendPinName = "";
            CapPinName = "";
            SendBitName = "";
            CapBitName = "";
            SendBitStr = "";
            CapBitStr = "";
            Subr = "";
            TimeSet = "";
            Payload = "";
        }

        public object Clone()
        {
            return MemberwiseClone();
        }
    }
}
