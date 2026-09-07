using System;
using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

namespace BinCutScriptLib
{
    internal static class HvInstanceDataBuilder
    {
        internal static void SetInstanceInfo(OneGradeSearch oneGradeSearch, List<List<PatternRow>> patternRows, ref SiteInfo[] siteInfoArray, int powerIdx, int searchStep, string curInstanceName)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].SiteIsBinOut || !siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                PowerZone pwrRef = siteInfoArray[site].AllPowers[powerIdx];
                int stepDiffTmp = pwrRef.Step - pwrRef.FinalStep;
                if (stepDiffTmp < 0)
                {
                    stepDiffTmp = Math.Abs(stepDiffTmp) - 1;
                }

                if (pwrRef.SearchStatus == EnumSearchStatus.Search)
                {
                    HandleStatusSearch(oneGradeSearch, patternRows, siteInfoArray, powerIdx, searchStep, curInstanceName, site, pwrRef, stepDiffTmp);
                }
                else
                {
                    HandleStatusOther(oneGradeSearch, patternRows, siteInfoArray, powerIdx, searchStep, curInstanceName, site, pwrRef, stepDiffTmp);
                }
            }
        }

        private static void HandleStatusOther(OneGradeSearch oneGradeSearch, List<List<PatternRow>> patternRows, SiteInfo[] siteInfoArray, int powerIdx, int searchStep, string curInstanceName, int site, PowerZone powerZone, int stepDiffTmp)
        {
            // 1. Initialize instance via unified helper with default values
            InstanceData oneInstanceData = CreateBaseInstanceData(oneGradeSearch, patternRows, siteInfoArray, powerIdx, searchStep, curInstanceName, site, powerZone, stepDiffTmp, 0.0, 0.0, -1, -1);

            // 2. Add directly to target collection
            siteInfoArray[site].InstanceList.Add(oneInstanceData);
        }

        private static void HandleStatusSearch(OneGradeSearch oneGradeSearch, List<List<PatternRow>> patternRows, SiteInfo[] siteInfoArray, int powerIdx, int searchStep, string curInstanceName, int site, PowerZone powerZone, int stepDiffTmp)
        {
            int bin = powerZone.GetFinalBin();

            // 1. Initialize instance via unified helper with active search metrics
            InstanceData oneInstanceData = CreateBaseInstanceData(oneGradeSearch, patternRows, siteInfoArray, powerIdx, searchStep, curInstanceName, site, powerZone, stepDiffTmp, powerZone.IdsValue, powerZone.GetFinalLvcc(), powerZone.GetFinalEqName(), bin);

            // 2. Apply explicit post-search state mutations
            powerZone.Step = powerZone.FinalStep;
            powerZone.Bin = bin;
            siteInfoArray[site].Bin = bin;

            // 3. Add to target collection
            siteInfoArray[site].InstanceList.Add(oneInstanceData);
        }

        private static InstanceData CreateBaseInstanceData(OneGradeSearch oneGradeSearch, List<List<PatternRow>> patternRows, SiteInfo[] siteInfoArray, int powerIdx, int searchStep, string curInstanceName, int site, PowerZone powerZone, int stepDiffTmp, double ids, double lvcc, int eqns, int bin)
        {
            var instanceData = new InstanceData
            {
                InstanceName = curInstanceName,
                PowersIdx = powerIdx,
                Ids = ids,
                Lvcc = lvcc,
                Eqns = eqns,
                Bin = bin,
                IsCheckPassByInstance = oneGradeSearch.InstanceBinCut!.IsBvPass,
                IsSearch = oneGradeSearch.InstanceBinCut.IsSearch,
                FinalStep = powerZone.FinalStep,
                UsedSteps = searchStep,
                ActualSteps = searchStep - (powerZone.SearchStatus != EnumSearchStatus.Search ? 0 : stepDiffTmp),
                PatternRows = [.. patternRows.Select(y => y.Where(x => x.Site == site).ToList())]
            };

            // Shared conditional parsing optimization
            if (BinCutConfig.IsDoAll)
            {
                foreach (PatternInfo pattern in siteInfoArray[site].PatResultList)
                {
                    if (pattern.IsFail)
                    {
                        instanceData.FailPatternData.Add(pattern);
                    }
                }
            }

            AlgorithmBaseHelpers.SetPatternRow(null, ref instanceData);

            return instanceData;
        }
    }
}
