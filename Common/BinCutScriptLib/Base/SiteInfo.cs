using System;
using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Static;

using CommonLib.Extension;

using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.PowerBinning;

namespace BinCutScriptLib.Base
{
    public class SiteInfo
    {
        public List<string> InFlowKey = [];

        public int TuchNum;
        //from "Site    X_Coord     Y_Coord"

        public int XCoor = -1;
        public int YCoor = -1;
        public int Site;
        private int _bin;
        public int Bin
        {
            get { return Math.Max(AllPowers.Count > 0 ? AllPowers.Max(x => x.Bin) : _bin, _bin); }
            set { _bin = value; }
        }
        //from "Site    Sort     Bin"
        public int Sort;
        //from "Site    Sort     Bin"
        public int SortBin;
        public List<int> HarvBin = [];
        public int SortLineNo = -1;
        public bool IsActiveSite;
        public bool IsIdsHotBinOut;
        public bool SiteIsBinOut
        {
            get { return AllPowers.Exists(x => x.IsBinOut) || IsIdsHotBinOut; }
        }

        public bool IsShutDown;

        #region powerbinning
        public int PowerBinningSeq;
        public PowerBinningSheet? PowerBinning;
        public List<Tuple<string, string>> PowerBinningPTotalSummary = [];
        public string PowerBinningFail = "N/A";
        public string PowerBinningBinXFail = "N/A";
        public Dictionary<string, double> IdsOnList = [];
        public Dictionary<string, double> IdsOffList = [];
        #endregion

        public string Currentdssc = "";
        public double CurrentdynamicOffset;
        //Get "real" IDS value at "IDS_SetWriteDecimal" lines
        public List<IdsData> RealIds = [];
        public List<PowerZone> AllPowers = [];
        public List<InstanceData> InstanceList = [];
        public readonly CheckResult CheckResult = new();
        public List<PatternInfo> PatternList = [];
        public List<PatternInfo> PatResultList = [];
        public bool IsTotalCurrCheckPass
        {
            get
            {
                return CheckResult.IsBvPass && CheckResult.IsInterpolationPass && CheckResult.IsPowerBinningPass
                     && CheckResult.IsDsscPass && CheckResult.IsLvResultPass && CheckResult.IsBinoutStatusPass;
            }
        }

        public Dictionary<string, double> EfuseProductValue = new(StringExtensions.IgnoreCase);

        #region CP2
        public List<EFuseRow> EFuseValues = [];
        #endregion

        public Dictionary<string, string> HarvesFlags = [];
        public bool IsBinCutConfig { get; set; }
        public bool IsPreVddSearch;
        public bool IsBin4Cand;
        public bool IsHarvFail;

        public SiteInfo() { }

        public SiteInfo(SiteInfo siteInfo)
        {
            if (siteInfo == null)
            {
                return;
            }

            InFlowKey = [.. siteInfo.InFlowKey];
            TuchNum = siteInfo.TuchNum;
            XCoor = siteInfo.XCoor;
            YCoor = siteInfo.YCoor;
            Site = siteInfo.Site;
            // Copy the backing field directly (not via the Bin property setter/getter,
            // which is computed from AllPowers rather than a plain passthrough).
            _bin = siteInfo._bin;
            Sort = siteInfo.Sort;
            SortBin = siteInfo.SortBin;
            HarvBin = [.. siteInfo.HarvBin];
            SortLineNo = siteInfo.SortLineNo;
            IsActiveSite = siteInfo.IsActiveSite;
            IsIdsHotBinOut = siteInfo.IsIdsHotBinOut;
            IsShutDown = siteInfo.IsShutDown;
            PowerBinningSeq = siteInfo.PowerBinningSeq;
            PowerBinning = siteInfo.PowerBinning?.Copy();
            PowerBinningPTotalSummary = [.. siteInfo.PowerBinningPTotalSummary];
            PowerBinningFail = siteInfo.PowerBinningFail;
            PowerBinningBinXFail = siteInfo.PowerBinningBinXFail;
            IdsOnList = new Dictionary<string, double>(siteInfo.IdsOnList);
            IdsOffList = new Dictionary<string, double>(siteInfo.IdsOffList);
            Currentdssc = siteInfo.Currentdssc;
            CurrentdynamicOffset = siteInfo.CurrentdynamicOffset;
            RealIds = [.. siteInfo.RealIds.Select(x => x.Copy())];
            AllPowers = [.. siteInfo.AllPowers.Select(x => x.Copy())];
            InstanceList = [.. siteInfo.InstanceList.Select(x => x.Copy())];
            CheckResult = siteInfo.CheckResult.Copy();
            PatternList = [.. siteInfo.PatternList.Select(x => x.Copy())];
            PatResultList = [.. siteInfo.PatResultList.Select(x => x.Copy())];
            EfuseProductValue = new Dictionary<string, double>(siteInfo.EfuseProductValue, StringExtensions.IgnoreCase);
            EFuseValues = [.. siteInfo.EFuseValues.Select(x => x.Copy())];
            HarvesFlags = new Dictionary<string, string>(siteInfo.HarvesFlags);
            IsBinCutConfig = siteInfo.IsBinCutConfig;
            IsPreVddSearch = siteInfo.IsPreVddSearch;
            IsBin4Cand = siteInfo.IsBin4Cand;
            IsHarvFail = siteInfo.IsHarvFail;
        }

        public SiteInfo Copy()
        {
            return new SiteInfo(this);
        }

        public PowerZone GetPowerZone(BinCutExpress binCutExpress)
        {
            var powerZone = new PowerZone();
            string powerName = BinCutAlgorithmService.GetPowerByName(binCutExpress.Express);
            int pPowerIdx = BinCutAlgorithmService.GetPowerIndex(powerName, this);
            if (pPowerIdx != -1)
            {
                powerZone = AllPowers[pPowerIdx];
            }

            return powerZone;
        }

        public void GetInherit(string parName, string sonName, out int allowEqlParIdx, out int allowEqlSonIdx)
        {
            allowEqlParIdx = -1;
            allowEqlSonIdx = -1;
            for (int pwrIdx = 0; pwrIdx < AllPowers.Count; pwrIdx++)
            {
                if (AllPowers[pwrIdx].PinMode.Contains(parName))
                {
                    allowEqlParIdx = pwrIdx;
                }

                if (AllPowers[pwrIdx].PinMode.Contains(sonName))
                {
                    allowEqlSonIdx = pwrIdx;
                }

                if (allowEqlParIdx != -1 && allowEqlSonIdx != -1)
                {
                    break;
                }
            }
        }

        public PowerStep GetFinalPassStep(List<BvName> bvNames)
        {
            for (int i = bvNames.Count - 1; i >= 0; i--)
            {
                string mode = bvNames[i].Mode;
                PowerZone powerZone = AllPowers.Find(x => x.Mode == mode)!;
                if (powerZone.SearchStatus == EnumSearchStatus.Search)
                {
                    return powerZone.PossibleSteps[powerZone.FinalStep];
                }

                PowerStep? finalPassStep = powerZone.AllSteps.Find(x => x.IsLastPatternPassForCof);
                if (finalPassStep != null)
                {
                    return finalPassStep;
                }
            }
            return null!;
        }

        public void ChangeFromIdsMode2Linear(int powerIdx)
        {
            AllPowers[powerIdx].IdsMode = EnumIdsMode.Linear;
            AllPowers[powerIdx].StopStep = AllPowers[powerIdx].GetPosCount() - 1;
            CheckResult.PatPassCnt = 0;
            CheckResult.PatFailCnt = 0;
        }

        public void InitDiceByPowers(List<PinInfo> pinInfos)
        {
            foreach (PinInfo pinInfo in pinInfos)
            {
                var powerZone = new PowerZone
                {
                    PinMode = pinInfo.PinMode,
                    Pin = pinInfo.Pin,
                    Mode = pinInfo.Mode,
                    Bin = Bin
                };
                AllPowers.Add(powerZone);
            }
        }

        public void InitHarvFlag()
        {
            List<string> flags = BinCutConfig.ProjectConfig.GetHarvFlags();
            if (flags == null)
            {
                return;
            }

            foreach (string flag in flags)
            {
                HarvesFlags.Add(flag, "F");
            }
        }

        public void SetHarvResult(string flag, bool status)
        {
            if (!HarvesFlags.ContainsKey(flag))
            {
                return;
            }

            HarvesFlags[flag] = status ? "T" : "F";
        }
    }
}
