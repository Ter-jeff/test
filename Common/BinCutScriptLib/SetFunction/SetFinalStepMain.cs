using System.Linq;

using BinCutScriptLib.Base;

namespace BinCutScriptLib.SetFunction
{
    public class SetFinalStepMain
    {
        public static void SetFinalStep(ref SiteInfo[] siteInfoArray, int powerIdx, EnumBcConfig enumBcConfig, string curInstanceName, string harvestBinningFlag)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (!siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                if (siteInfoArray[site].AllPowers[powerIdx].IsLastPatPass == EnumPatPf.Pass)
                {
                    siteInfoArray[site].CheckResult.PatPassCnt++;
                }

                if (siteInfoArray[site].AllPowers[powerIdx].IsLastPatPass == EnumPatPf.Fail)
                {
                    siteInfoArray[site].CheckResult.PatFailCnt++;
                }

                var patResultList = siteInfoArray[site].PatResultList.OrderBy(x => x.Bin).ThenByDescending(y => y.EqName).ThenBy(z => z.StepNum).ToList();
                PowerZone pwr = siteInfoArray[site].AllPowers[powerIdx];
                PatternInfo? foundLastFailPat = patResultList.FindLast(x => x.IsFail);
                siteInfoArray[site].AllPowers[powerIdx].SearchStatus = EnumSearchStatus.Search;
                if (harvestBinningFlag.Length != 0 && siteInfoArray[site].HarvesFlags.ContainsKey(harvestBinningFlag))
                {
                    siteInfoArray[site].SetHarvResult(harvestBinningFlag, false);
                }

                if (foundLastFailPat == null)
                {
                    // All pass case
                    siteInfoArray[site].AllPowers[powerIdx].FinalStep = ConvPwrStep(pwr, patResultList.First());
                    continue;
                }

                int lastFailStep = ConvPwrStep(pwr, foundLastFailPat);
                if (ComparePattern(foundLastFailPat, patResultList.Last()))
                {
                    int maxStepNum = patResultList.FindAll(x => x.Bin.Equals(patResultList.Last().Bin) && x.EqName.Equals(patResultList.Last().EqName)).Max(y => y.StepNum);
                    if (patResultList.FindAll(x => x.Bin.Equals(patResultList.Last().Bin) && x.EqName.Equals(patResultList.Last().EqName) && x.StepNum.Equals(maxStepNum) && x.IsFail).Count != 0)
                    {
                        siteInfoArray[site].AllPowers[powerIdx].SearchStatus = EnumSearchStatus.BinOut;
                        if (harvestBinningFlag.Length != 0 && siteInfoArray[site].HarvesFlags.ContainsKey(harvestBinningFlag))
                        {
                            siteInfoArray[site].SetHarvResult(harvestBinningFlag, true);
                        }
                    }
                    else
                    {
                        // special case if failstep is highest step and last search pattern results are pass , no binout.
                        siteInfoArray[site].AllPowers[powerIdx].FinalStep = lastFailStep;
                    }
                }
                else
                {
                    // normal case
                    siteInfoArray[site].AllPowers[powerIdx].FinalStep = lastFailStep + 1;
                    if (ComparePattern(foundLastFailPat, patResultList.First()))
                    {
                        // Only check lowest step in first search steps pattern result
                        IGrouping<int, PatternInfo> firstSearchStep = patResultList.FindAll(x => x.Bin.Equals(patResultList.First().Bin) && x.EqName.Equals(patResultList.First().EqName)).GroupBy(z => z.StepNum).First();
                        foreach (PatternInfo pattern in firstSearchStep)
                        {
                            if (pattern.IsFail)
                            {
                                // if pattern fail, select the lowest step +1
                                siteInfoArray[site].AllPowers[powerIdx].FinalStep = lastFailStep + 1;
                                break;
                            }
                            // if pattern pass ,select the lowest step
                            siteInfoArray[site].AllPowers[powerIdx].FinalStep = lastFailStep;
                        }
                    }
                }
            }
        }

        public static int ConvPwrStep(PowerZone powerZone, PatternInfo patternInfo)
        {
            return powerZone.PossibleSteps.FindIndex(x => x.EqName.Equals(patternInfo.EqName) && x.Bin.Equals(patternInfo.Bin));
        }

        public static bool ComparePattern(PatternInfo pattern1, PatternInfo pattern2)
        {
            return pattern1.Bin.Equals(pattern2.Bin) && pattern1.EqName.Equals(pattern2.EqName);
        }
    }
}
