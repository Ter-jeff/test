using System.Collections.Generic;

using BinCutScriptLib.Base;

using TestPlanLib.BinCut;

namespace BinCutScriptLib.SetFunction.SetStartStep
{
    public class IdsDistribution : InheritanceBase
    {
        public static void SetAllPowers(List<int> testedSiteInMode, ref SiteInfo[] siteInfoArray, BvName bvName, IdsDistributionTable idsDistributionTable, string curInstance, bool isConditionMatched)
        {
            int instType = GetInstType(curInstance);
            string powerName = bvName.Name;
            int powerIdx = bvName.Index;
            if (idsDistributionTable != null && idsDistributionTable.AllIdsPowers.Count != 0 && isConditionMatched)
            {
                for (int site = 0; site < siteInfoArray.Length; site++)
                {
                    if (siteInfoArray[site].AllPowers.Count == 0 || !siteInfoArray[site].IsActiveSite)
                    {
                        continue;
                    }

                    if (testedSiteInMode.Exists(x => x == site))
                    {
                        continue;
                    }

                    if (siteInfoArray[site].AllPowers[powerIdx].IsInterplation)
                    {
                        continue;
                    }

                    if (siteInfoArray[site].IsPreVddSearch)
                    {
                        continue;
                    }

                    double idsValue = siteInfoArray[site].AllPowers[powerIdx].IdsValue;
                    IdsPower? idsPower = idsDistributionTable.AllIdsPowers.Find(x => powerName.Contains(x.PowerName));
                    if (idsPower == null)
                    {
                        return;
                    }

                    IdsInfo idsInfo = idsPower.IdsInfos[instType];
                    MoveStepByIds(siteInfoArray, idsInfo, idsValue, site, powerIdx);
                }
            }
        }

        private static void MoveStepByIds(SiteInfo[] siteInfoArray, IdsInfo idsInfo, double idsValue, int site, int powerIdx)
        {
            for (int index = 0; index < idsInfo.IdsRng.Count - 1; index++)
            {
                if (idsValue >= idsInfo.IdsRng[index] && idsValue < idsInfo.IdsRng[index + 1])
                {
                    int startEq = idsInfo.StartBin[index];
                    if (siteInfoArray[site].IsBin4Cand)
                    {
                        continue;
                    }

                    for (int stepIdx = 0; stepIdx < siteInfoArray[site].AllPowers[powerIdx].GetPosCount(); stepIdx++)
                    {
                        if (siteInfoArray[site].AllPowers[powerIdx].PossibleSteps[stepIdx].EqName == startEq) //start equation
                        {
                            siteInfoArray[site].AllPowers[powerIdx].Step = stepIdx;
                            siteInfoArray[site].AllPowers[powerIdx].StartStep = stepIdx;
                            siteInfoArray[site].AllPowers[powerIdx].FinalStep = stepIdx;
                            siteInfoArray[site].Bin = siteInfoArray[site].AllPowers[powerIdx].GetFinalBin();
                            if (stepIdx != 0) //IDS
                            {
                                siteInfoArray[site].AllPowers[powerIdx].IdsMode = EnumIdsMode.Ids;
                                siteInfoArray[site].AllPowers[powerIdx].StopStep = 0;
                            }
                            else
                            {
                                siteInfoArray[site].AllPowers[powerIdx].IdsMode = EnumIdsMode.Linear;
                                siteInfoArray[site].AllPowers[powerIdx].StopStep = siteInfoArray[site].AllPowers[powerIdx].GetPosCount() - 1;
                            }
                            break;
                        }
                    }
                    break;
                }
            }
        }

        private static int GetInstType(string curInstance)
        {
            //TD		MBIST		SPI		TMPS		LDCBFD
            int instType;
            if (curInstance.Contains("TD", System.StringComparison.OrdinalIgnoreCase))
            {
                instType = 0;
            }
            else if (curInstance.Contains("MBIST", System.StringComparison.OrdinalIgnoreCase))
            {
                instType = 1;
            }
            else if (curInstance.Contains("SPI", System.StringComparison.OrdinalIgnoreCase) || curInstance.Contains("RTOS", System.StringComparison.OrdinalIgnoreCase))
            {
                instType = 2;
            }
            else
            {
                //<-所有未分類的皆視為TD
                instType = 0;
            }

            return instType;
        }

        public static void ChangeFromIdsMode2LinearBeforeNextInstance(SiteInfo[] siteInfoArray, int powerIdx)
        {
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (siteInfoArray[site].SiteIsBinOut || !siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                if (siteInfoArray[site].AllPowers[powerIdx].IdsMode == EnumIdsMode.Ids)
                {
                    siteInfoArray[site].ChangeFromIdsMode2Linear(powerIdx);
                }
            }
        }
    }
}
