using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Comparer.PowerBinning;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.SetFunction;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;

using PinInfo = TestPlanLib.BinCut.Binning.PinInfo;

namespace BinCutScriptLib
{
    public class BinCutCheckMain : AlgorithmBase
    {
        public List<AllowEqualBase> AllowEqualList = [];
        public List<Tuple<string, string>> SkipPwrList = [];
        public List<PinInfo> PinInfos;
        public List<string> PowerNames;
        public InheritanceManager InheritanceList = new();

        public BinCutCheckMain(BinCutScriptMain binCutScriptMain, List<PinInfo> pinInfos, List<string> powerPins, BinningTables binningTables, BinCutFlowTables binCutFlowTables)
            : base(binCutScriptMain)
        {
            PinInfos = [.. pinInfos.Where(x => powerPins.Exists(y => y.EqualsIgnoreCase(x.Pin)))];
            PowerNames = [.. PinInfos.Select(x => x.PinMode)];
            AllowEqualList = binningTables.First().GetAllowEqualList();
            InheritanceList.Read(binningTables.First().GetInherits(PowerNames));
            SkipPwrList = binCutFlowTables.GetSkipPwrList();
            PinInfos = pinInfos;
        }

        protected override void Run(string dataLogFile, int maxSite)
        {
            using StreamReader sr = DataLogReader.GetStreamReader(dataLogFile);
            using var sw = new StreamWriter(OutputFile);
            while (!sr.EndOfStream)
            {
                OneTouchDown = new OneTouchDown();
                var sites = new SiteInfo[maxSite];
                var allDiceBackup = new SiteInfo[maxSite];
                bool cofMode = BinCutConfig.FlagUseCofInstance;
                bool hasPreSearch = false;
                var curBinDic = new Dictionary<int, int>();
                for (int site = 0; site < maxSite; site++)
                {
                    sites[site] = new SiteInfo { Site = -1, Sort = -1, SortBin = -1, XCoor = -1, YCoor = -1, Bin = 1, HarvBin = [], TuchNum = TchCnt };
                    sites[site].InitDiceByPowers(PinInfos);
                    sites[site].InitHarvFlag();
                }

                if (!BinCutConfig.FlagSyncUpDcvsOutputEnable)
                {
                    BinCutPrint.PrintErrorMessage(sw, "SyncUp DCVS Output is DISABLED");
                }

                if (!GetOneTouchDown(sr, ref OneTouchDown, ref sites))
                {
                    break;
                }

                OneTouchDown.GetHarvestConfigForEachTouchDown(sites);

                PrintCurrentSite(sites);

                //Get Multi pin result
                Dictionary<string, Dictionary<string, bool>> multiPinResult = OneTouchDown.GetMultiPinResult();

                //Get flagstate
                Dictionary<string, Dictionary<string, bool>> flagState = OneTouchDown.GetFlagState();

                //Get real ids value 
                bool isFoundSetWriteDecimal = OneTouchDown.GetRealIds(sw, ref sites);
                OneTouchDown.GetIdsOnAndOff(ref sites);
                OneTouchDown.GetEfuseIds(sw, ref sites, PowerNames);
                AlgorithmBaseHelpers.SetHarvResult(sites, flagState);
                AlgorithmBaseHelpers.SetHarvResult(sites, multiPinResult);

                //Get IDS Value and create all powerNames (if powerNames.count = 0 mean IDS fail) 
                JudgeStoredIdsMain.CheckJudgeStoredIds(sw, Job.JobType, ref OneTouchDown, ref sites, isFoundSetWriteDecimal);
                //Must have => Product_Identifier & product

                CfgFuseMain.GetCfgFuse(ref OneTouchDown, ref sites, Job.JobType);

                ProAction(sites, PinInfos, sw);

                if (BinCutConfig.FlagVddbinHarvestBin4RunBin1Eqn1)
                {
                    //Set initial step to Bin1 Eq 1 when BIN4_CAND flag is true (by Site)
                    foreach (SiteInfo allDic in sites)
                    {
                        if (!allDic.HarvesFlags.TryGetValue("BIN4_CAND", out string? flag))
                        {
                            continue;
                        }

                        if (flag == "T")
                        {
                            AlgorithmBaseHelpers.SetEqFromBin1Eq1(allDic);
                            allDic.IsBin4Cand = true;
                        }
                    }
                }
                //Set initial step to Bin X when F_Bin4X_Pass flag is true (by Site)
                AlgorithmBaseHelpers.SetBinForFBin4XPass(ref sites, multiPinResult);

                LvSearchMain.HandleLvVbt(sw, ref sites, ref allDiceBackup, cofMode, ref hasPreSearch, ref curBinDic, out OneGradeSearch oneGradeSearch, out _, OneTouchDown, CheckManager, CurInstanceName, BvNames, BvNamesBackup, Job, PowerNames, Method, InheritanceList, AllowEqualList, TchCnt);

                #region Recheck eqNLines when skip test is true
                if (oneGradeSearch.VBinResultLines.Count > 0 || OneTouchDown.EnRows.Count > 0)
                {
                    SkipModeMain.BeSkipReArrangePowerTable(sw, ref sites, oneGradeSearch.VBinResultLines, OneTouchDown.EnRows, InheritanceList, AllowEqualList, SkipPwrList, Job.JobType);
                }
                #endregion

                AdjustVddBinningMain.ComparePre_Adjust_VddBinningVbt(sw, ref sites, OneTouchDown, InheritanceList, AllowEqualList);

                #region Power_Binning_Calculation
                var powerBinningComparesor = new PowerBinningComparesor(OneTouchDown, CheckManager, Job.JobType);
                powerBinningComparesor.CheckPowerBinning(ref sites, sw);
                #endregion

                #region Compare <PrintOutVddBinning>
                if (OneTouchDown.GetPrintOutVddBinning(out List<BinCutLineBase> printOutLines))
                {
                    PrintOutVddBinningComparer.CompareCp1PrintOut(sw, ref sites, ref printOutLines, Job.JobType);
                }
                #endregion

                AdjustVddBinningMain.CompareAdjust_VddBinningVbt(sw, ref sites, OneTouchDown, InheritanceList, AllowEqualList, Job.JobType);

                sites = HvSearchMain.HandleHvVbt(sw, sites, Job, OneTouchDown, CurInstanceName, CheckManager, TchCnt, out oneGradeSearch, out bool isSearch, PowerNames);

                PostSearchMain.HandlePostVbt(sw, ref sites, ref oneGradeSearch, OneTouchDown, Job, CurInstanceName, CheckManager, TchCnt, PowerNames, ref isSearch);

                CurrentDiceInfos.AddRange(AlgorithmBaseHelpers.AddDiceInfo(sites));

                BinCutPrint.PrintMissingLines(sw, CheckManager.MissingBv);
                TchCnt++;
            }
        }
    }
}
