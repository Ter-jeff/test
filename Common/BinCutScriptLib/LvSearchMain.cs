using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm;
using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.SetFunction;
using BinCutScriptLib.SetFunction.SetStartStep;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using TestPlanLib;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib
{
    internal class LvSearchMain
    {
        private static void LvSearchVbt(OneGradeSearch oneGradeSearch, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, bool cofMode, List<BvName> bvNames, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, string curInstanceName, Job job, CheckManager checkManager, int tchCnt, List<string> powerNames)
        {
            if (oneGradeSearch.Steps.Count == 0 && oneGradeSearch.SearchResultWoTestLine.Count == 0)
            {
                return;
            }

            if (oneGradeSearch.SearchResultWoTestLine.Count != 0)
            {
                SearchResultWoTestLine1.SetSearchResultbySearchResultWoTestLine(ref siteInfoArray, oneGradeSearch.SearchResultWoTestLine);
            }
            BvName bvName = oneGradeSearch.GetFirstBvName(powerNames);
            if (bvName == null)
            {
                return;
            }

            if (!bvNames.Exists(x => x.Name.EqualsIgnoreCase(bvName.Name)))
            {
                inheritanceManager.SetInheritModeEnable(bvName.Mode);
            }

            SetStartStepMain.WorkFlow(inheritanceManager, allowEqualBases, enumBcConfig, curInstanceName, streamWriter, ref siteInfoArray, oneGradeSearch.EnRows, bvName, bvNames);

            BinCutCheckMainHelpers.AddBvName(bvName, ref bvNames);
            var preStepDssc = new List<BinCutLineBase>();
            var prePatternRows = new List<PatternRow>();
            int searchStep = 0;
            var patternRows = new List<List<PatternRow>>();
            bool hasHarvFlag = oneGradeSearch.HarvestBinningFlag.Length != 0;

            HandleOneLvGrade(oneGradeSearch, streamWriter, ref siteInfoArray, cofMode, enumBcConfig, curInstanceName, job, checkManager, tchCnt, bvName, ref preStepDssc, ref prePatternRows, ref searchStep, patternRows);

            IdsDistribution.ChangeFromIdsMode2LinearBeforeNextInstance(siteInfoArray, bvName.Index);

            AlgorithmBaseHelpers.SetInstanceInfoLv(curInstanceName, patternRows, ref siteInfoArray, bvName.Index, searchStep, cofMode, oneGradeSearch);

            //Compare IDS/LVCC/EQN/PASSBIN results
            var lvResultsComparer = new LvResultsComparer(curInstanceName);
            lvResultsComparer.CompareLvResults(streamWriter, ref siteInfoArray, oneGradeSearch.JudgeResultLines, bvName, job.JobType, oneGradeSearch.IsHarvAssignedTrue);
            if (cofMode)
            {
                lvResultsComparer.ComparePlResults(streamWriter, ref siteInfoArray, oneGradeSearch.PatternResultLines);
            }

            #region Vddbin_COF_StepInheritance to modify step when bin out
            if (enumBcConfig == EnumBcConfig.Vddbin_COF_StepInheritance || hasHarvFlag || oneGradeSearch.IsHarvAssignedTrue || enumBcConfig == EnumBcConfig.Vddbin_COF_StepInheritance_New_Logic)
            {
                Inheritance.ModifyStepWhenBinout(siteInfoArray, bvName, bvNames, oneGradeSearch.IsHarvAssignedTrue, hasHarvFlag);
            }
            #endregion

            PostAction(oneGradeSearch, siteInfoArray, cofMode);
        }

        private static void PostAction(OneGradeSearch oneGradeSearch, SiteInfo[] siteInfoArray, bool cofMode)
        {
            if (cofMode)
            {
                foreach (SiteInfo diceInfo in siteInfoArray)
                {
                    diceInfo.PatternList.Clear();
                }

                oneGradeSearch.PatternResultLines.Clear();
            }

            foreach (SiteInfo diceInfo in siteInfoArray)
            {
                diceInfo.PatResultList.Clear();
            }
        }

        private static void HandleOneLvGrade(OneGradeSearch oneGradeSearch, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, bool cofMode, EnumBcConfig enumBcConfig, string curInstanceName, Job job, CheckManager checkManager, int tchCnt, BvName bvName, ref List<BinCutLineBase> binCutLineBases, ref List<PatternRow> patternRows, ref int searchStep, List<List<PatternRow>> patternRowList)
        {
            #region OneLvGrade
            for (int i = 0; i < oneGradeSearch.Steps.Count; i++)
            {
                OneStep step = oneGradeSearch.Steps[i];
                bool isInitSkipMoveStep = new BvCompareMain(job.JobType, checkManager, tchCnt, oneGradeSearch).DataCompare(streamWriter, ref siteInfoArray, bvName, ref step, i, ref binCutLineBases, ref patternRows);

                AlgorithmBaseHelpers.AddCofPattern(cofMode, isInitSkipMoveStep, oneGradeSearch, siteInfoArray, bvName);

                JudgePassFailMain.IsStepPf(ref siteInfoArray, ref step, bvName.Index, searchStep, isInitSkipMoveStep);

                new HarvestSourceCodeManager(checkManager).Check(streamWriter, ref siteInfoArray, step, curInstanceName);

                if (isInitSkipMoveStep)
                {
                    continue;
                }

                SetFinalStepMain.SetFinalStep(ref siteInfoArray, bvName.Index, enumBcConfig, curInstanceName, oneGradeSearch.HarvestBinningFlag);

                patternRowList.Add([.. step.OneStepPatternRows.Select(x => x.Copy())]);

                bool isSchEnd = CheckSearchEndMain.CheckSearchEnd(ref siteInfoArray, bvName.Index);

                if (!isSchEnd)
                {
                    IncreaseStepMain.IncreaseStep(ref siteInfoArray, bvName);
                }

                searchStep++;
            }
            #endregion
        }

        internal static void HandleLvVbt(StreamWriter streamWriter, ref SiteInfo[] allDice, ref SiteInfo[] allDiceBackup, bool cofMode, ref bool hasPreSearch, ref Dictionary<int, int> curBinDic, out OneGradeSearch oneGradeSearch, out bool isSearch, OneTouchDown oneTouchDown, CheckManager checkManager, string curInstanceName, List<BvName> bvNames, List<BvName> bvNamesBackup, Job job, List<string> powerNames, EnumBcConfig enumBcConfig, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, int tchCnt)
        {
            #region LV SECTION
            checkManager.TotalBvCnt += oneTouchDown.GetBvCount();
            checkManager.TotalHarvestSourceCodeCnt += oneTouchDown.GetHarvestSourceCodeCntCount();
            checkManager.TotalDsscCnt += oneTouchDown.GetDsscCnt();
            checkManager.MissingBv = oneTouchDown.GetMissingLine();
            bvNames.Clear();
            isSearch = false;
            bool isPreSearch = false;
            var preSearchList = new List<string>();
            while (new GetCp1LvTests(job.JobType, streamWriter).GetInstances(ref oneTouchDown, out oneGradeSearch, ref curInstanceName, out isSearch))
            {
                foreach (SiteInfo diceInfo in allDice)
                {
                    if (diceInfo.IsShutDown &&
                        oneGradeSearch.Steps.Exists(
                            x => x.OneStepBvLineInfo.Exists(y => y.Site == diceInfo.Site)))
                    {
                        BinCutPrint.PrintSiteMismatchErrorMessage(streamWriter, curInstanceName, diceInfo.Site);
                        diceInfo.IsShutDown = false;
                    }
                }

                //Backup for FSTP, PreVddSearch
                if (oneGradeSearch.FstpStatus == EnumFstpStatus.Start || oneGradeSearch.PreVddStatus == EnumPreVddStatus.Start)
                {
                    BackupDiceInfo(ref allDice, ref allDiceBackup, true, ref curBinDic, bvNames, bvNamesBackup);
                    //For special search from bin1 eq1
                    if (oneGradeSearch.PreVddStatus == EnumPreVddStatus.Start)
                    {
                        isPreSearch = true;
                        allDice.ToList().ForEach(x => x.IsPreVddSearch = true);
                    }
                }
                if (isPreSearch)
                {
                    string mode = curInstanceName.Split('_').First();
                    if (!preSearchList.Exists(x => x == mode))
                    {
                        preSearchList.Add(mode);
                        foreach (SiteInfo diceInfo in allDice.Where(diceinfo => diceinfo.AllPowers.Count != 0))
                        {
                            AlgorithmBaseHelpers.SetEqFromBin1Eq1ByMode(diceInfo, mode);
                        }
                    }
                }
                //Restore for normal search
                if (oneGradeSearch.FstpStatus == EnumFstpStatus.Stop || oneGradeSearch.PreVddStatus == EnumPreVddStatus.Stop)
                {
                    if (oneGradeSearch.PreVddStatus == EnumPreVddStatus.Stop)
                    {
                        allDice.ToList().ForEach(x => x.IsPreVddSearch = false);
                        hasPreSearch = true;
                    }
                    else
                    {
                        BackupDiceInfo(ref allDice, ref allDiceBackup, false, ref curBinDic, bvNames, bvNamesBackup);
                    }
                    if (isPreSearch)
                    {
                        isPreSearch = false;
                    }
                }

                if (hasPreSearch)
                {
                    if (oneGradeSearch.IsReJudge)
                    {
                        BackupDiceInfo(ref allDice, ref allDiceBackup, false, ref curBinDic, bvNames, bvNamesBackup);
                        foreach (KeyValuePair<int, int> item in curBinDic)
                        {
                            int site = item.Key;
                            int bin = item.Value;
                            SiteInfo curSite = allDice[site];
                            if (curSite.Bin != bin)
                            {
                                AlgorithmBaseHelpers.SetTargetBin(curSite, bin);
                            }
                        }
                    }
                    hasPreSearch = false;
                }

                AlgorithmBaseHelpers.SetAllDiceByLog(allDice, oneTouchDown, oneGradeSearch);

                if (isSearch)
                {
                    LvSearchVbt(oneGradeSearch, streamWriter, ref allDice, cofMode, bvNames, inheritanceManager, allowEqualBases, enumBcConfig, curInstanceName, job, checkManager, tchCnt, powerNames);
                }
                else
                {
                    HvSearchMain.HvSearchVbt(oneGradeSearch, streamWriter, ref allDice, job, checkManager, tchCnt, curInstanceName, powerNames);
                }

                if (BinCutConfig.FlagSyncUpDcvsOutputEnable && !oneGradeSearch.Steps.Any(x => x.OneStepSyncUpLine.Count > 0) && oneGradeSearch.Steps.Any(x => x.OneStepBvLineInfo.Any(y => y.InstType == "TD")))
                {
                    BinCutPrint.PrintSyncUpErrorMessage(curInstanceName, streamWriter);
                }
            }
            #endregion
        }

        public static void BackupDiceInfo(ref SiteInfo[] allDice, ref SiteInfo[] allDicesBackup, bool isBackup, ref Dictionary<int, int> curBinDic, List<BvName> bvNames, List<BvName> bvNamesBackup)
        {
            if (isBackup)
            {
                allDicesBackup = [.. allDice.Select(x => x.Copy())];
                bvNamesBackup = [.. bvNames.Select(x => x.Copy())];
            }
            else
            {
                curBinDic = allDice.Where(x => x.Site != -1).ToDictionary(x => x.Site, x => x.Bin);
                allDice = allDicesBackup;
                bvNames = bvNamesBackup;
            }
        }
    }
}
