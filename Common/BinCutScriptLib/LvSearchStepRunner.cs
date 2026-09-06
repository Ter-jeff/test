using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.SetFunction;

using TestPlanLib;

namespace BinCutScriptLib
{
    internal static class LvSearchStepRunner
    {
        internal static void RunStep(Job job, CheckManager checkManager, int tchCnt, OneGradeSearch oneGradeSearch, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, BvName bvName, OneStep oneStep, ref List<BinCutLineBase> binCutLineBases, ref List<PatternRow> patternRows, string curInstanceName, bool cofMode, EnumBcConfig enumBcConfig, List<List<PatternRow>> patternRowList, ref int searchStep)
        {
            bool isInitSkipMoveStep = new BvCompareMain(job.JobType, checkManager, tchCnt, oneGradeSearch).DataCompareCs(streamWriter, ref siteInfoArray, bvName, ref oneStep, ref binCutLineBases, ref patternRows, curInstanceName);

            AlgorithmBaseHelpers.AddCofPattern(cofMode, isInitSkipMoveStep, oneGradeSearch, siteInfoArray, bvName);

            JudgePassFailMain.IsStepPf(ref siteInfoArray, ref oneStep, bvName.Index, searchStep, isInitSkipMoveStep);

            new HarvestSourceCodeManager(checkManager).CheckHarvestSourceCodeCs(streamWriter, ref siteInfoArray, oneStep.OneHarvestSourceCodetRows);

            if (isInitSkipMoveStep)
            {
                return;
            }

            SetFinalStepMain.SetFinalStep(ref siteInfoArray, bvName.Index, enumBcConfig, curInstanceName, oneGradeSearch.HarvestBinningFlag);

            patternRowList.Add([.. oneStep.OneStepPatternRows.Select(x => x.Copy())]);

            bool isSchEnd = CheckSearchEndMain.CheckSearchEnd(ref siteInfoArray, bvName.Index);

            if (!isSchEnd)
            {
                IncreaseStepMain.IncreaseStep(ref siteInfoArray, bvName);
            }

            searchStep++;
        }
    }
}
