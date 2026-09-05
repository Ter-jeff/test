using System.Collections.Generic;

using CommonLib.Datalog;
using CommonLib.Extension;

using EfuseCheckCmdLib.DataStructure;

namespace EfuseCheckCmdLib.Datalog
{
    public class DiceInfo
    {
        //from "Site    X_Coord     Y_Coord"
        public int XCoor;
        public int YCoor;
        public int Site;
        //from "Site    Sort     Bin"
        public int Sort;
        //from "Site    Sort     Bin"
        public int SortBin;
        public string PrrLotId = "";
        public string PrrWaferId = "";
        public string PrrX = "";
        public string PrrY = "";
        public string PrrCode = "";

        public string? EFuseLotNumber = null;
        public string? EFuseWaferId = null;
        public string? EFuseDieX = null;
        public string? EFuseDieY = null;
        public string? HramEcid53Bit = null;
        public string? SvmCFuse288Bits = null;
        public string Scenario = "";
        public string DramType = "";
        public string CurrentJobStage = "";
        public int JobNum = -1;
        public List<EfuseRow> EfuseRows = [];
        public Dictionary<string, EfuseDatalogItem> Items = new(StringExtensions.IgnoreCase);
        public List<PrrRow> PrrRows = [];
        public List<SetWriteVariableLine> SetWriteVariableLines = [];
        public List<LimitRow> LimitRows { get; internal set; } = [];
    }
}
