using System.Collections.Generic;

namespace Cautogen.common.ReaderWriter.Reader.InputDataBase
{
    public class HardIpSeqInfo
    {
        public string SeqName { get; set; }
        public string PinList { get; set; }
        public string SortPinList { get; set; }
        public List<string> MeasIPowerPinList = new List<string>();
        public List<string> MeasVPowerPinList = new List<string>();
        public List<string> MeasIIoPinList = new List<string>();
        public List<string> MeasVIoPinList = new List<string>();
        public List<string> MeasFPinList = new List<string>();
        public List<string> MeasVdiffPinList = new List<string>();
        public List<string> MeasIdiffPinList = new List<string>();
        public List<string> MeasVdiff2PinList = new List<string>();
        public List<string> MeasFdiffPinList = new List<string>();
        public List<string> MeasVocmPinList = new List<string>();
        public List<string> MeasZPinList = new List<string>();
        public List<string> MeasRPinList = new List<string>();

        public HardIpSeqInfo()
        {
            SortPinList = "";
            PinList = "";
            SeqName = "";
        }
    }


    public class HardIpReference
    {
        public int CapBit { get; set; }
        public int SendBit { get; set; }
        public string Payload { get; set; }
        public string Version { get; set; }
        public string TimeSet { get; set; }
        public string Subr { get; set; }
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
        public string MeasRStr { get; set; }
        public string MeasSeqStr { get; set; }
        public string MeasZStr { get; set; }
        public string DigSrcAssign { get; set; }
        public string DigSrcDataWidth { get; set; }
        public List<string> MeasIPowerPinList = new List<string>();
        public List<string> MeasVPowerPinList = new List<string>();
        public List<string> MeasIIoPinList = new List<string>();
        public List<string> MeasVIoPinList = new List<string>();
        public List<string> MeasFPinList = new List<string>();
        public List<string> MeasVdiffPinList = new List<string>();
        public List<string> MeasIdiffPinList = new List<string>();
        public List<string> MeasVdiff2PinList = new List<string>();
        public List<string> MeasFdiffPinList = new List<string>();
        public List<string> MeasVocmPinList = new List<string>();
        public List<string> MeasZPinList = new List<string>();
        public List<string> MeasRPinList = new List<string>();
        public List<HardIpSeqInfo> SeqInfo = new List<HardIpSeqInfo>();
        public bool IllegalChar = false;

        public HardIpReference()
        {
            DigSrcDataWidth = "";
            DigSrcAssign = "";
            MeasZStr = "";
            MeasSeqStr = "";
            MeasRStr = "";
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
            Version = "";
        }
    }
}
