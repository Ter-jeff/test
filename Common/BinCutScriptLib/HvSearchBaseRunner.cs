using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Comparer.BVComparer;
using BinCutScriptLib.SetFunction;

using TestPlanLib;

namespace BinCutScriptLib
{
    internal static class HvSearchBaseRunner
    {
        internal static SiteInfo[] RunSearchBase(OneGradeSearch oneGradeSearch, StreamWriter streamWriter, SiteInfo[] siteInfoArray, Job job, CheckManager checkManager, int tchCnt, string curInstanceName, BvName bvName)
        {
            var patternRows = new List<List<PatternRow>>();
            for (int i = 0; i < oneGradeSearch.Steps.Count; i++)
            {
                OneStep step = oneGradeSearch.Steps[i];
                var comparer = new BvComparerNew(oneGradeSearch, job.JobType, checkManager, tchCnt, EnumSearchType.NonGradeSearch, EnumPrintType.LV);
                comparer.CompareBvString(streamWriter, ref siteInfoArray, i, step, bvName);

                new HarvestSourceCodeManager(checkManager).Check(streamWriter, ref siteInfoArray, step, curInstanceName);

                JudgePassFailMain.IsPatPfWithoutBinout(streamWriter, ref siteInfoArray, step.OneStepPatternRows, bvName, step.OneStepNoBinOut, curInstanceName, step.OneStepMfstpNoBinOut);
                patternRows.Add([.. step.OneStepPatternRows.Select(x => x.Copy())]);

                var datalogPatternRows = oneGradeSearch.Steps.SelectMany(x => x.OneStepPatternRows.Select(y => y.Copy()).ToList()).ToList();
                HvInstanceDataBuilder.SetInstanceInfo(oneGradeSearch, [datalogPatternRows], ref siteInfoArray, bvName.Index, 1, curInstanceName);
            }

            return siteInfoArray;
        }
    }
}
