using System.Collections.Generic;

using BinCutScriptLib.Base;

namespace BinCutScriptLib.SetFunction
{
    internal class IncreaseStepMain
    {
        public static void IncreaseStep(ref SiteInfo[] siteInfoArray, BvName bvName)
        {
            int powerIdx = bvName.Index;
            var idsStepchangeList = new List<int>();
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (!siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                PowerZone pwrRef = siteInfoArray[site].AllPowers[powerIdx];
                if (siteInfoArray[site].AllPowers[powerIdx].IdsMode == EnumIdsMode.Linear)
                {
                    if (siteInfoArray[site].AllPowers[powerIdx].Step != siteInfoArray[site].AllPowers[powerIdx].StopStep)
                    {
                        siteInfoArray[site].AllPowers[powerIdx].Step++;
                    }
                    //else
                    //{
                    //    if(BinCutConfig.IsDoAll)
                    //    {
                    //        allDice[site].AllPowers[powerIdx].Step = 0;
                    //    }
                    //}
                }
                else if (siteInfoArray[site].AllPowers[powerIdx].IdsMode == EnumIdsMode.Ids)
                {
                    if (siteInfoArray[site].AllPowers[powerIdx].Step != siteInfoArray[site].AllPowers[powerIdx].StopStep)
                    {
                        int step = siteInfoArray[site].AllPowers[powerIdx].Step;
                        if (step == 0)
                        {
                            continue;
                        }

                        siteInfoArray[site].AllPowers[powerIdx].Step--;
                        idsStepchangeList.Add(site);
                    }
                }
                siteInfoArray[site].Bin = pwrRef.PossibleSteps[pwrRef.Step].Bin;
            }
        }
    }
}
