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
using BinCutScriptLib.SetFunction.SetStartStep;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;

using PinInfo = TestPlanLib.BinCut.Binning.PinInfo;

namespace BinCutScriptLib
{
    public class BinCutCheckMainCs : AlgorithmBase
    {
        public List<AllowEqualBase> AllowEqualList = [];
        public List<Tuple<string, string>> SkipPwrList = [];
        public List<PinInfo> PinInfos;
        public List<string> PowerNames;
        public InheritanceManager InheritanceList = new();

        public BinCutCheckMainCs(BinCutScriptMain binCutScriptMain, List<PinInfo> pinInfos, List<string> powerPins, BinningTables binningTables, BinCutFlowTables binCutFlowTables)
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
                bool cofMode = BinCutConfig.DebugBinCutCofStored;
                SiteInfo[] sites = InitSite(maxSite, PinInfos);

                if (!GetOneTouchDown(sr, ref OneTouchDown, ref sites))
                {
                    break;
                }

                OneTouchDown.GetHarvestConfigForEachTouchDownCs(sites, Job.JobType);

                PrintCurrentSite(sites);

                //Get real ids value
                bool isFoundSetWriteDecimal = OneTouchDown.GetRealIdsCs(sw, ref sites);

                //Get efuse ids value
                OneTouchDown.GetEfuseIdsCs(sw, ref sites, PowerNames);

                //Get IDS Value and create all powerNames (if powerNames.count = 0 mean IDS fail) 
                JudgeStoredIdsMain.CheckJudgeStoredIdsCs(sw, Job.JobType, ref OneTouchDown, ref sites, isFoundSetWriteDecimal, PinInfos);
                //Must have => Product_Identifier & product

                CfgFuseMain.GetCfgFuse(ref OneTouchDown, ref sites, Job.JobType);
                //Must have => Product_Identifier & product

                ReadDvfmManager.ReadProduct(sw, Job.JobType, ref OneTouchDown, ref sites);
                //For C# Product value (Only Debug print)

                AssignFuseCsReader.GetAssignFuseCs(ref OneTouchDown, ref sites, Job.JobType, PinInfos);

                ProAction(sites, PinInfos, sw);

                List<BinCutLineBase> lines = AlgorithmBaseHelpers.GotoVddbinningstart(ref OneTouchDown);
                if (lines.Count != 0)
                {
                    OneGradeSearch oneGradeSearch = new OneGradeSearch();
                    CheckManager.TotalBvCnt += OneTouchDown.GetBvCountCs();
                    CheckManager.MissingBv = OneTouchDown.GetMissingLineCsharp();

                    BvNames.Clear();

                    oneGradeSearch = LvSearchMainCs.HandleLvCs(sw, cofMode, ref sites, Job, OneTouchDown, CurInstanceName, SkipPwrList, PinInfos, BvNames, InheritanceList, AllowEqualList, Method, CheckManager, TchCnt);

                    AlgorithmBaseHelpers.SyncUpStepForAffiliatePin(sites);

                    #region Interpolation
                    if (OneTouchDown.EnRows.Count > 0)
                    {
                        var skips = OneTouchDown.EnRows.Where(x => x.SkipTest).ToList();
                        Interpolation.SetAllPowersCs(sw, ref sites, skips, InheritanceList);
                    }
                    #endregion

                    #region Skip Test
                    if (oneGradeSearch.VBinResultLines.Count > 0)
                    {
                        List<Tuple<string, string>> skipList = SkipPwrList.FindAll(x => x.Item1.EqualsIgnoreCase(Job.JobType.ToString()));
                        SkipModeMain.BeSkipReArrangePowerTableCs(sw, ref sites, oneGradeSearch.VBinResultLines, InheritanceList, AllowEqualList, skipList);
                    }
                    #endregion

                    #region Power_Binning
                    var powerBinningComparesor = new PowerBinningComparesorCs(OneTouchDown, CheckManager, Job.JobType);
                    powerBinningComparesor.CheckPowerBinningCs(ref sites, sw);
                    #endregion

                    AdjustVddBinningMain.ComparePre_Adjust_VddBinningCs(sw, ref sites, OneTouchDown, InheritanceList, AllowEqualList);

                    #region Compare <PrintOutVddBinning>
                    if (OneTouchDown.GetPrintOutVddBinningCs(out List<BinCutLineBase> printOutLines))
                    {
                        var printOutVddComparesor = new PrintOutVddBinningComparer();
                        PrintOutVddBinningComparer.CompareCp1PrintOutCs(sw, ref sites, ref printOutLines);
                    }
                    #endregion

                    AdjustVddBinningMain.CompareAdjust_VddBinningCs(sw, ref sites, OneTouchDown, InheritanceList, AllowEqualList, Job.JobType);

                    oneGradeSearch = HvSearchMain.HandleHvCs(sw, ref sites, Job, OneTouchDown, CurInstanceName, CheckManager, TchCnt, PinInfos);

                    PostSearchMain.HandlePostCs(sw, ref sites, ref oneGradeSearch, OneTouchDown, Job, CurInstanceName, CheckManager, TchCnt, PinInfos);

                    CurrentDiceInfos.AddRange(AlgorithmBaseHelpers.AddDiceInfo(sites));

                    BinCutPrint.PrintMissingLines(sw, CheckManager.MissingBv);
                }
                TchCnt++;
            }
        }
    }
}
