using System.Collections.Generic;

using BinCutScriptLib.Base.Line;

namespace BinCutScriptLib.Comparer
{
    public class CheckManager
    {
        public List<BinCutLineBase> MissingBv = [];
        public int TotalBvCnt;
        public int TotalBvCheckCnt;
        public int TotalDsscCnt;
        public int TotalDsscCheckCnt;
        public int TotalPowerBinningCnt;
        public int TotalPowerBinningCheckCnt;
        public int TotalHarvestSourceCodeCnt;
        public int TotalHarvestSourceCodeCheckCnt;

        public void Initialize()
        {
            TotalBvCnt = 0;
            TotalBvCheckCnt = 0;
            TotalDsscCnt = 0;
            TotalDsscCheckCnt = 0;
            TotalPowerBinningCnt = 0;
            TotalPowerBinningCheckCnt = 0;
            TotalHarvestSourceCodeCnt = 0;
            TotalHarvestSourceCodeCheckCnt = 0;
        }

        public void Add(CheckManager checkManager)
        {
            TotalBvCnt += checkManager.TotalBvCnt;
            TotalBvCheckCnt += checkManager.TotalBvCheckCnt;
            TotalDsscCnt += checkManager.TotalDsscCnt;
            TotalDsscCheckCnt += checkManager.TotalDsscCheckCnt;
            TotalPowerBinningCnt += checkManager.TotalPowerBinningCnt;
            TotalPowerBinningCheckCnt += checkManager.TotalPowerBinningCheckCnt;
            TotalHarvestSourceCodeCnt += checkManager.TotalHarvestSourceCodeCnt;
            TotalHarvestSourceCodeCheckCnt += checkManager.TotalHarvestSourceCodeCheckCnt;
        }

        public void CheckBvLine(BinCutLineBase binCutLineBase)
        {
            TotalBvCheckCnt++;
            if (MissingBv.Exists(x => x.LineNo.Equals(binCutLineBase.LineNo)))
            {
                MissingBv.Remove(MissingBv.Find(x => x.LineNo.Equals(binCutLineBase.LineNo))!);
            }
        }

        public void CheckHarvestSourceCodeLine(BinCutLineBase binCutLineBase)
        {
            TotalHarvestSourceCodeCheckCnt++;
            if (MissingBv.Exists(x => x.LineNo.Equals(binCutLineBase.LineNo)))
            {
                MissingBv.Remove(MissingBv.Find(x => x.LineNo.Equals(binCutLineBase.LineNo))!);
            }
        }

        public void CheckDsscLine(BinCutLineBase binCutLineBase)
        {
            TotalDsscCheckCnt++;
            if (MissingBv.Exists(x => x.LineNo.Equals(binCutLineBase.LineNo)))
            {
                MissingBv.Remove(MissingBv.Find(x => x.LineNo.Equals(binCutLineBase.LineNo))!);
            }
        }

        public void CheckPowerBinningLine(BinCutLineBase binCutLineBase)
        {
            TotalPowerBinningCheckCnt++;
            if (MissingBv.Exists(x => x.LineNo.Equals(binCutLineBase.LineNo)))
            {
                MissingBv.Remove(MissingBv.Find(x => x.LineNo.Equals(binCutLineBase.LineNo))!);
            }
        }
    }
}
