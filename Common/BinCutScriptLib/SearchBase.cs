using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm;
using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Comparer.BVComparer;
using BinCutScriptLib.Printer;
using BinCutScriptLib.SetFunction;
using BinCutScriptLib.Static;

using TestPlanLib;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib
{
    internal class SearchBase
    {
        protected static void EvaluateSyncAndPatterns(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, OneGradeSearch oneGradeSearch, string instanceName, BvName bvName, bool forceSyncCheck)
        {
            if (oneGradeSearch?.Steps == null)
            {
                return;
            }

            // Validate synchronization rules safely across execution pipelines
            bool missingSyncLines = !oneGradeSearch.Steps.Any(x => x.OneStepSyncUpLine != null && x.OneStepSyncUpLine.Count > 0);
            bool containsTdElements = oneGradeSearch.Steps.Any(x => x.OneStepBvLineInfo != null && x.OneStepBvLineInfo.Any(y => y?.InstType == "TD"));

            if (forceSyncCheck && missingSyncLines && containsTdElements)
            {
                BinCutPrint.PrintSyncUpErrorMessage(instanceName, streamWriter);
            }

            // Evaluate overall pattern passes/failures without binout
            var collectedPatterns = oneGradeSearch.Steps.SelectMany(x => x.OneStepPatternRows ?? Enumerable.Empty<PatternRow>()).ToList();
            bool hasNoBinOut = oneGradeSearch.Steps.Any(x => x.OneStepNoBinOut);

            JudgePassFailMain.IsPatPfWithoutBinout(streamWriter, ref siteInfoArray, collectedPatterns, bvName, hasNoBinOut, instanceName);
        }

        protected static void SyncTouchDownBins1(SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown)
        {
            if (siteInfoArray == null || oneTouchDown?.Bins == null)
            {
                return;
            }

            foreach (KeyValuePair<int, int> item in oneTouchDown.Bins)
            {
                if (item.Key >= 0 && item.Key < siteInfoArray.Length && siteInfoArray[item.Key] != null)
                {
                    siteInfoArray[item.Key].Bin = item.Value;
                }
            }
        }

        protected static void ValidateSiteMismatches(StreamWriter streamWriter, SiteInfo[] siteInfoArray, OneGradeSearch oneGradeSearch, string instanceName, bool isCsMode)
        {
            if (siteInfoArray == null || oneGradeSearch?.Steps == null)
            {
                return;
            }

            foreach (SiteInfo diceInfo in siteInfoArray)
            {
                if (diceInfo != null && diceInfo.IsShutDown && oneGradeSearch.Steps.Any(x => x.OneStepBvLineInfo != null && x.OneStepBvLineInfo.Any(y => y.Site == diceInfo.Site)))
                {
                    if (!isCsMode || !(BinCutConfig.DebugBinCutCofStored || BinCutConfig.IsDoAll))
                    {
                        BinCutPrint.PrintSiteMismatchErrorMessage(streamWriter, instanceName, diceInfo.Site);
                    }
                    diceInfo.IsShutDown = false;
                }
            }
        }

        protected static OneGradeSearch ProcessHvPostCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, Job job, ref OneTouchDown oneTouchDown, ref string curInstanceName, CheckManager checkManager, int tchCnt, List<PinInfo> pinInfos, EnumPrintType enumPrintType, SearchLine searchLine)
        {
            OneGradeSearch oneGradeSearch;
            while (searchLine.GetInstancesCs(ref oneTouchDown, out oneGradeSearch, ref curInstanceName, out _))
            {
                // 1. Process shared structural state updates
                ValidateSiteMismatches(streamWriter, siteInfoArray, oneGradeSearch, curInstanceName, isCsMode: true);
                SyncTouchDownBins1(siteInfoArray, oneTouchDown);

                BvName bvName = oneGradeSearch.GetFirstBvNameCs(pinInfos);
                if (bvName == null)
                {
                    continue;
                }

                // 2. Step Engine Loop execution
                for (int i = 0; i < oneGradeSearch.Steps.Count; i++)
                {
                    OneStep step = oneGradeSearch.Steps[i];
                    var cp1Comparer = new BvComparerNew(oneGradeSearch, job.JobType, checkManager, tchCnt, EnumSearchType.NonGradeSearch, enumPrintType);
                    cp1Comparer.CompareBvStringCs(streamWriter, ref siteInfoArray, step, bvName, curInstanceName);

                    new HarvestSourceCodeManager(checkManager).CheckHarvestSourceCodeCs(streamWriter, ref siteInfoArray, step.OneHarvestSourceCodetRows);
                }

                // 3. Process trailing tracking logs and data persistence
                EvaluateSyncAndPatterns(streamWriter, ref siteInfoArray, oneGradeSearch, curInstanceName, bvName, forceSyncCheck: true);
            }

            return oneGradeSearch;
        }
    }
}
