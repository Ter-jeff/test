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
    internal class PostSearchMain : SearchBase
    {
        internal static void HandlePostCs(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, ref OneGradeSearch oneGradeSearch, OneTouchDown oneTouchDown, Job job, string curInstanceName, CheckManager checkManager, int tchCnt, List<PinInfo> pinInfos)
        {
            if (oneTouchDown.Lines.Count == 0)
            {
                return;
            }
            EnumPrintType printType = EnumPrintType.POST;

            var testFetcher = new GetCp1PostTests(job.JobType, streamWriter);
            oneGradeSearch = ProcessHvPostCs(streamWriter, ref siteInfoArray, job, ref oneTouchDown, ref curInstanceName, checkManager, tchCnt, pinInfos, printType, testFetcher);
        }

        internal static void HandlePostVbt(StreamWriter streamWriter, ref SiteInfo[] siteInfoArray, ref OneGradeSearch oneGradeSearch, OneTouchDown oneTouchDown, Job job, string curInstanceName, CheckManager checkManager, int tchCnt, List<string> powerNames, ref bool isSearch)
        {
            #region POST
            if (oneTouchDown.Lines.Count > 0)
            {
                while (new GetCp1PostTests(job.JobType, streamWriter).GetInstances(ref oneTouchDown, out oneGradeSearch, ref curInstanceName, out isSearch))
                {
                    foreach (SiteInfo diceInfo in siteInfoArray)
                    {
                        if (diceInfo.IsShutDown &&
                            oneGradeSearch.Steps.Exists(
                                x => x.OneStepBvLineInfo.Exists(y => y.Site == diceInfo.Site)))
                        {
                            BinCutPrint.PrintSiteMismatchErrorMessage(streamWriter, curInstanceName, diceInfo.Site);
                            diceInfo.IsShutDown = false;
                        }
                    }
                    foreach (KeyValuePair<int, int> item in oneTouchDown.Bins)
                    {
                        siteInfoArray[item.Key].Bin = item.Value;
                    }

                    BvName bvName = oneGradeSearch.GetFirstBvName(powerNames);
                    if (bvName == null)
                    {
                        continue;
                    }

                    for (int i = 0; i < oneGradeSearch.Steps.Count; i++)
                    {
                        OneStep step = oneGradeSearch.Steps[i];
                        //var cp1Comparer = new Cp1BvPostComparer(oneGradeSearch, Job, CheckManager, TchCnt, EnumSearchType.NonGradeSearch);
                        var cp1Comparer = new BvComparerNew(oneGradeSearch, job.JobType, checkManager, tchCnt, EnumSearchType.NonGradeSearch, EnumPrintType.POST);
                        cp1Comparer.CompareBvString(streamWriter, ref siteInfoArray, i, step, bvName);
                        new HarvestSourceCodeManager(checkManager).Check(streamWriter, ref siteInfoArray, step, curInstanceName);
                    }
                    if (BinCutConfig.FlagSyncUpDcvsOutputEnable && !oneGradeSearch.Steps.Any(x => x.OneStepSyncUpLine.Count > 0) && oneGradeSearch.Steps.Any(x => x.OneStepBvLineInfo.Any(y => y.InstType == "TD")))
                    {
                        BinCutPrint.PrintSyncUpErrorMessage(curInstanceName, streamWriter);
                    }

                    JudgePassFailMain.IsPatPfWithoutBinout(streamWriter, ref siteInfoArray, [.. oneGradeSearch.Steps.SelectMany(x => x.OneStepPatternRows)], bvName, oneGradeSearch.Steps.Exists(x => x.OneStepNoBinOut), curInstanceName);
                    List<PatternRow> patternRows = [.. oneGradeSearch.Steps.SelectMany(x => x.OneStepPatternRows.Select(y => y.Copy()).ToList())];
                    HvSearchMain.SetInstanceInfoHv(oneGradeSearch, [patternRows], ref siteInfoArray, bvName.Index, 1, curInstanceName);
                }
            }
            #endregion
        }
    }
}
