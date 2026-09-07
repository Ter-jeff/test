using System;
using System.Collections.Generic;
using System.IO;

using BinCutScriptLib.Algorithm;
using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Comparer;
using BinCutScriptLib.Printer;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib
{
    internal class LvSearchMainCs
    {
        internal static OneGradeSearch HandleLvCs(StreamWriter streamWriter, bool cofMode, ref SiteInfo[] siteInfoArray, Job job, OneTouchDown oneTouchDown, string curInstanceName, List<Tuple<string, string>> skipPwrList, List<PinInfo> pinInfos, List<BvName> bvNames, InheritanceManager inheritanceManager, List<AllowEqualBase> allowEqualBases, EnumBcConfig enumBcConfig, CheckManager checkManager, int tchCnt)
        {
            OneGradeSearch oneGradeSearch;
            #region LV SECTION
            while (new GetCp1LvTests(job.JobType, streamWriter).GetInstancesCs(ref oneTouchDown, out oneGradeSearch, ref curInstanceName, out bool _))
            {
                foreach (SiteInfo diceInfo in siteInfoArray)
                {
                    if (diceInfo.IsShutDown && oneGradeSearch.Steps.Exists(x => x.OneStepBvLineInfo.Exists(y => y.Site == diceInfo.Site)))
                    {
                        if (!(BinCutConfig.DebugBinCutCofStored || BinCutConfig.IsDoAll))
                        {
                            BinCutPrint.PrintSiteMismatchErrorMessage(streamWriter, curInstanceName, diceInfo.Site);
                        }
                        diceInfo.IsShutDown = false;
                    }
                }

                AlgorithmBaseHelpers.SetAllDiceByLog(siteInfoArray, oneTouchDown, oneGradeSearch);

                if (job.JobType == EnumJob.CP1)
                {
                    List<Tuple<string, string>> skipList = skipPwrList.FindAll(x => x.Item1.EqualsIgnoreCase(job.JobType.ToString()));
                    LvSearchMainCs1.LvSearchCs(oneGradeSearch, streamWriter, ref siteInfoArray, cofMode, bvNames, inheritanceManager, allowEqualBases, enumBcConfig, curInstanceName, job, checkManager, tchCnt, skipList, pinInfos);
                }
                else
                {
                    HvSearchMain.HvSearchCs(oneGradeSearch, streamWriter, ref siteInfoArray, job, checkManager, tchCnt, curInstanceName, pinInfos);
                }
            }
            #endregion
            return oneGradeSearch;
        }
    }
}
