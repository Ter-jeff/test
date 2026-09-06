using System;
using System.Collections.Generic;
using System.IO;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.SetFunction.SetStartStep;

using CommonLib.Extension;

using TestPlanLib;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib
{
    internal class LvSearchMainCs1
    {
        internal static void LvSearchCs(OneGradeSearch oneGradeSearch, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, bool cofMode, List<BvName> bvNames, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, string curInstanceName, Job job, CheckManager checkManager, int tchCnt, List<Tuple<string, string>> skipPwrList, List<PinInfo> pinInfos)
        {
            if (oneGradeSearch.Steps.Count == 0 && oneGradeSearch.SearchResultWoTestLine.Count == 0)
            {
                return;
            }

            if (oneGradeSearch.SearchResultWoTestLine.Count != 0)
            {
                SearchResultWoTestLine1.SetSearchResultbySearchResultWoTestLine(ref siteInfoArray, oneGradeSearch.SearchResultWoTestLine);
            }
            BvName bvName = oneGradeSearch.GetFirstBvNameCs(pinInfos);
            if (bvName == null)
            {
                return;
            }

            if (!bvNames.Exists(x => x.Mode.EqualsIgnoreCase(bvName.Mode)))
            {
                inheritanceManager.SetInheritModeEnable(bvName.Mode);
            }

            SetStartStepMain.WorkFlowCs(inheritanceManager, allowEqualBases, enumBcConfig, curInstanceName, streamWriter, ref siteInfoArray, oneGradeSearch.EnRows, oneGradeSearch.VBinResultLines, bvName, bvNames, skipPwrList);

            BinCutCheckMainHelpers.AddBvName(bvName, ref bvNames);
            var preStepDssc = new List<BinCutLineBase>();
            var prePatternRows = new List<PatternRow>();
            int searchStep = 0;
            var patternRows = new List<List<PatternRow>>();
            bool hasHarvFlag = oneGradeSearch.HarvestBinningFlag.Length != 0;

            #region OneLvGrade
            for (int i = 0; i < oneGradeSearch.Steps.Count; i++)
            {
                OneStep step = oneGradeSearch.Steps[i];
                LvSearchStepRunner.RunStep(job, checkManager, tchCnt, oneGradeSearch, streamWriter, ref siteInfoArray, bvName, step, ref preStepDssc, ref prePatternRows, curInstanceName, cofMode, enumBcConfig, patternRows, ref searchStep);
            }
            #endregion

            IdsDistribution.ChangeFromIdsMode2LinearBeforeNextInstance(siteInfoArray, bvName.Index);

            AlgorithmBaseHelpers.SetInstanceInfoLv(curInstanceName, patternRows, ref siteInfoArray, bvName.Index, searchStep, cofMode, oneGradeSearch);

            //Compare IDS/LVCC/EQN/PASSBIN results
            var lvResultsComparer = new LvResultsComparer(curInstanceName);
            lvResultsComparer.CompareLvResults(streamWriter, ref siteInfoArray, oneGradeSearch.JudgeResultLines, bvName, job.JobType, oneGradeSearch.IsHarvAssignedTrue);
            if (cofMode)
            {
                lvResultsComparer.ComparePlResults(streamWriter, ref siteInfoArray, oneGradeSearch.PatternResultLines);
            }

            //Csharp is from Vddbin_COF_StepInheritance to Debug_BinCutCOF_Stored
            if (hasHarvFlag || oneGradeSearch.IsHarvAssignedTrue || enumBcConfig == EnumBcConfig.Debug_BinCutCOF_Stored)
            {
                Inheritance.ModifyStepWhenBinout(siteInfoArray, bvName, bvNames, oneGradeSearch.IsHarvAssignedTrue, hasHarvFlag);
            }

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
    }
}
