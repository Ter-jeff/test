using System.IO;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Comparer.BVComparer;

using TestPlanLib;

namespace BinCutScriptLib
{
    internal static class HvVbtStepRunner
    {
        internal static void RunVbtStep(OneGradeSearch oneGradeSearch, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, Job job, CheckManager checkManager, int tchCnt, int stepIndex, OneStep oneStep, BvName bvName, string curInstanceName, EnumPrintType enumPrintType)
        {
            var cp1Comparer = new BvComparerNew(oneGradeSearch, job.JobType, checkManager, tchCnt, EnumSearchType.NonGradeSearch, enumPrintType);
            cp1Comparer.CompareBvString(streamWriter, ref siteInfoArray, stepIndex, oneStep, bvName);

            new HarvestSourceCodeManager(checkManager).Check(streamWriter, ref siteInfoArray, oneStep, curInstanceName);
        }
    }
}
