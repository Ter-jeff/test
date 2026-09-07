using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Algorithm;
using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using TestPlanLib;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib
{
    internal class HvSearchMain : SearchBase
    {
        internal static OneGradeSearch HandleHvCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, Job job, OneTouchDown oneTouchDown, string curInstanceName, CheckManager checkManager, int tchCnt, List<PinInfo> pinInfos)
        {
            OneGradeSearch oneGradeSearch;
            EnumPrintType printType = curInstanceName.ContainsIgnoreCase("_POST_") ? EnumPrintType.POST : EnumPrintType.HV;

            var testFetcher = new GetCp1HvTests(job.JobType, streamWriter);
            oneGradeSearch = ProcessHvPostCs(streamWriter, ref siteInfoArray, job, ref oneTouchDown, ref curInstanceName, checkManager, tchCnt, pinInfos, printType, testFetcher);

            return oneGradeSearch;
        }

        internal static SiteInfo[] HandleHvVbt(StreamWriter streamWriter, SiteInfo[] siteInfoArray, Job job, OneTouchDown oneTouchDown, string curInstanceName, CheckManager checkManager, int tchCnt, out OneGradeSearch oneGradeSearch, out bool isSearch, List<string> powerNames)
        {
            var testFetcher = new GetCp1HvTests(job.JobType, streamWriter);

            while (testFetcher.GetInstances(ref oneTouchDown, out oneGradeSearch, ref curInstanceName, out isSearch))
            {
                // 1. Process shared structural state updates
                ValidateSiteMismatches(streamWriter, siteInfoArray, oneGradeSearch, curInstanceName, isCsMode: false);
                SyncTouchDownBins(siteInfoArray, oneTouchDown);

                BvName bvName = oneGradeSearch.GetFirstBvName(powerNames);
                if (bvName == null)
                {
                    continue;
                }

                // 2. Step Engine Loop execution
                for (int i = 0; i < oneGradeSearch.Steps.Count; i++)
                {
                    OneStep step = oneGradeSearch.Steps[i];
                    EnumPrintType printType = curInstanceName.ContainsIgnoreCase("_POST_") ? EnumPrintType.POST : EnumPrintType.HV;

                    HvVbtStepRunner.RunVbtStep(oneGradeSearch, streamWriter, ref siteInfoArray, job, checkManager, tchCnt, i, step, bvName, curInstanceName, printType);
                }

                // 3. Process trailing tracking logs and data persistence
                EvaluateSyncAndPatterns(streamWriter, ref siteInfoArray, oneGradeSearch, curInstanceName, bvName, forceSyncCheck: BinCutConfig.FlagSyncUpDcvsOutputEnable);

                List<PatternRow> patternRows = [.. oneGradeSearch.Steps.SelectMany(x => x.OneStepPatternRows.Select(y => y.Copy()))];
                HvInstanceDataBuilder.SetInstanceInfo(oneGradeSearch, [patternRows], ref siteInfoArray, bvName.Index, 1, curInstanceName);
            }

            return siteInfoArray;
        }

        #region Private Shared Modular Helpers

        private static void SyncTouchDownBins(SiteInfo[] siteInfoArray, OneTouchDown oneTouchDown)
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
        #endregion

        internal static void SetInstanceInfoHv(OneGradeSearch oneGradeSearch, List<List<PatternRow>> patternRows, ref SiteInfo[] siteInfoArray, int powerIdx, int searchStep, string curInstanceName)
        {
            HvInstanceDataBuilder.SetInstanceInfo(oneGradeSearch, patternRows, ref siteInfoArray, powerIdx, searchStep, curInstanceName);
        }

        internal static void HvSearchVbt(OneGradeSearch oneGradeSearch, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, Job job, CheckManager checkManager, int tchCnt, string curInstanceName, List<string> powerNames)
        {
            BvName bvName = oneGradeSearch.GetFirstBvName(powerNames);
            if (bvName == null)
            {
                return;
            }

            siteInfoArray = HvSearchBaseRunner.RunSearchBase(oneGradeSearch, streamWriter, siteInfoArray, job, checkManager, tchCnt, curInstanceName, bvName);
        }

        internal static void HvSearchCs(OneGradeSearch oneGradeSearch, StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, Job job, CheckManager checkManager, int tchCnt, string curInstanceName, List<PinInfo> pinInfos)
        {
            BvName bvName = oneGradeSearch.GetFirstBvNameCs(pinInfos);
            if (bvName == null)
            {
                return;
            }

            siteInfoArray = HvSearchBaseRunnerCs.RunSearchBase(oneGradeSearch, streamWriter, siteInfoArray, job, checkManager, tchCnt, curInstanceName, bvName);
        }
    }
}
