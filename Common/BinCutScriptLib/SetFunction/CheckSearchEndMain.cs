using BinCutScriptLib.Base;

namespace BinCutScriptLib.SetFunction
{
    internal class CheckSearchEndMain
    {
        public static bool CheckSearchEnd(ref SiteInfo[] siteInfoArray, int powerIdx)
        {
            bool isSearchEnd = true;
            for (int site = 0; site < siteInfoArray.Length; site++)
            {
                if (!siteInfoArray[site].IsActiveSite || siteInfoArray[site].AllPowers.Count == 0)
                {
                    continue;
                }

                if (siteInfoArray[site].AllPowers[powerIdx].IdsMode == EnumIdsMode.Linear &&
                    siteInfoArray[site].AllPowers[powerIdx].IsLastPatPass == EnumPatPf.Fail)
                {
                    isSearchEnd = false;
                }
                else if (siteInfoArray[site].AllPowers[powerIdx].IdsMode == EnumIdsMode.Ids)
                {
                    if (siteInfoArray[site].AllPowers[powerIdx].IsLastPatPass == EnumPatPf.Fail)
                    {
                        if (siteInfoArray[site].AllPowers[powerIdx].Step == siteInfoArray[site].AllPowers[powerIdx].StartStep)
                        {
                            siteInfoArray[site].ChangeFromIdsMode2Linear(powerIdx);
                            isSearchEnd = false;
                        }
                    }

                    if (siteInfoArray[site].AllPowers[powerIdx].IsLastPatPass == EnumPatPf.Pass)
                    {
                        isSearchEnd = false;
                    }
                }
            }
            return isSearchEnd;
        }
    }
}
