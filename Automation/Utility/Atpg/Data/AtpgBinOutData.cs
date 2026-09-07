using System.Collections.Generic;

namespace Automation.Utility.Atpg.Data
{
    public class AtpgBinOutData
    {
        public string Domain { get; }
        public string Block { get; }
        public string BinOp { get; }
        public string BinOutStage { get; }
        public bool IsByPassBinOut { get; }
        public bool IsHvLvBinOut { get; }
        public bool IsNeedEvsDeferredBinOut { get; }
        public List<string> FlagList { get; }

        public AtpgBinOutData(string domain, string block, string binOp, string binOutStage, bool isByPassBinOut, bool isHvLvBinOut, bool isNeedEvsDeferredBinOut, List<string> flagList)
        {
            Domain = domain;
            Block = block;
            BinOp = binOp;
            BinOutStage = BinOutStage;
            IsByPassBinOut = isByPassBinOut;
            IsHvLvBinOut = isHvLvBinOut;
            IsNeedEvsDeferredBinOut = isNeedEvsDeferredBinOut;
            FlagList = flagList;
        }
    }
}
