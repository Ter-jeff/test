using System.Collections.Generic;

using EfuseCheckCmdLib.IgxlLogLib;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    public class XDiceInfo
    {
        //from "Site    X_Coord     Y_Coord"
        public int XCoor = -1;
        public int YCoor = -1;
        public int Site = -1;
        //from "Site    Sort     Bin"
        public int Sort = -1;
        //real bin, read from data log
        //Result of die if exist PostCheck Test Instance
        public int SortBin = -1;
        //from "Site    Sort     Bin"
        public bool PostCheck;
        //Get "real" IDS value at "IDS_SetWriteDecimal" lines

        public string PrrLotId = "";
        public string PrrWaferId = "";
        public string PrrX = "";
        public string PrrY = "";
        public string PrrCode = "";
        #region IEDA Data
        public string? EFuseLotNumber;
        public string? EFuseWaferId;
        public string? EFuseDieX;
        public string? EFuseDieY;
        public string? HramEcid53Bit;
        public string? SvmCFuse288Bits;
        #endregion

        //only use by getLogBasicInfo(), the purpose is to get all performance power name from <Judge_stored_IDS>
        public List<string> PowerNames = [];
        public DeviceInfo Prober = new();
        //use by efuse check
        public List<EfuseDatalogItem> AllReadFromDssc = [];
        public List<EfuseDatalogItem> AllUdrVer = [];
        public Dictionary<string, string> IdsFuseInfo = [];
        public Dictionary<string, SetWriteItem> HardIpFuseInfo = [];
        public Dictionary<string, string> BvFuseInfo = [];
        public List<DataFormatDataRow> IdsMeasLines = [];
        public List<SetHipDicItem> HipMeasLines = [];
        public List<DataFormatDataRow> FusePatternLines = [];
        public FuseMpData? FuseMpDataSet;
        public Dictionary<string, EccData> EccInfo = [];
        public List<CheckSum> ChecksumList = [];
    }
}
