using System.Collections.Generic;
using System.IO;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Comparer.BVComparer;
using BinCutScriptLib.Static;

using IgxlLib.Enums;

namespace BinCutScriptLib.SetFunction
{
    internal class BvCompareMain
    {
        private readonly EnumJob _job;
        private readonly CheckManager _checkManager;
        private readonly int _tchCnt;
        private readonly BvComparerNew _bvComparerNew;

        public BvCompareMain(EnumJob enumJob, CheckManager checkManager, int tchCnt, OneGradeSearch oneGradeSearch)
        {
            _job = enumJob;
            _checkManager = checkManager;
            _tchCnt = tchCnt;
            _bvComparerNew = new BvComparerNew(oneGradeSearch, _job, _checkManager, _tchCnt, EnumSearchType.GradeSearch, EnumPrintType.LV);
        }

        public bool DataCompare(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, BvName bvName, ref OneStep oneStep, int stepIndex, ref List<BinCutLineBase> binCutLineBases, ref List<PatternRow> patternRows)
        {
            return DataCompareBase(ref siteInfoArray, ref oneStep, ref binCutLineBases, ref patternRows,
                (ref SiteInfo[] sites, OneStep step) =>
                {
                    return _bvComparerNew.CompareBvString(streamWriter, ref sites, stepIndex, step, bvName);
                });
        }

        public bool DataCompareCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, BvName bvName, ref OneStep oneStep, ref List<BinCutLineBase> binCutLineBases, ref List<PatternRow> patternRows, string curInstanceName)
        {
            // Capture curInstanceName configuration parameters inside localized closure scopes
            return DataCompareBase(ref siteInfoArray, ref oneStep, ref binCutLineBases, ref patternRows,
                (ref SiteInfo[] sites, OneStep step) =>
                {
                    return _bvComparerNew.CompareBvStringCs(streamWriter, ref sites, step, bvName, curInstanceName);
                });
        }

        private static bool DataCompareBase(ref SiteInfo[] siteInfoArray, ref OneStep oneStep, ref List<BinCutLineBase> binCutLineBases, ref List<PatternRow> patternRows, CompareAction compareAction)
        {
            if (binCutLineBases.Count != 0)
            {
                oneStep.OneStepDssc.AddRange(binCutLineBases);
                binCutLineBases.Clear();
            }

            if (patternRows.Count != 0)
            {
                //Add Nv pat into next pattern
                oneStep.OneStepPatternRows.InsertRange(0, patternRows);
                patternRows.Clear();
            }

            if (compareAction(ref siteInfoArray, oneStep))
            {
                if (oneStep.OneStepDssc.Count != 0)
                {
                    binCutLineBases.AddRange(oneStep.OneStepDssc);
                }

                if (oneStep.OneStepPatternRows.Count != 0)
                {
                    patternRows.AddRange(oneStep.OneStepPatternRows);
                }

                if (BinCutConfig.FlagUseCofInstance.Equals(true))
                {
                }
                return true;
            }

            return false;
        }

        private delegate bool CompareAction(ref SiteInfo[] sites, OneStep step);
    }
}
