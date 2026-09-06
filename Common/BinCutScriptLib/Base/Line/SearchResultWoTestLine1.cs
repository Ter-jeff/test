using System;
using System.Collections.Generic;
using System.Linq;

namespace BinCutScriptLib.Base.Line
{
    public class SearchResultWoTestLine1 : BinCutLineBase
    {
        private ResultLine GetSearchResultLine()
        {
            var searchResultLine = new ResultLine();
            string[] spilt = Line.Split([' ', ',', '.'], StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in spilt)
            {
                if (s.Contains("SITE", StringComparison.OrdinalIgnoreCase))
                {
                    _ = int.TryParse(s.Split(':').Last(), out searchResultLine.Site);
                    continue;
                }
                if (s.Contains("VDD_", StringComparison.OrdinalIgnoreCase))
                {
                    searchResultLine.Pmode = s;
                    continue;
                }
                if (s.Contains("EQN", StringComparison.OrdinalIgnoreCase))
                {
                    _ = int.TryParse(s.Split('=').Last(), out searchResultLine.Eqn);
                    continue;
                }
                if (s.Contains("GRADEVDD", StringComparison.OrdinalIgnoreCase))
                {
                    _ = int.TryParse(s.Split('=').Last(), out searchResultLine.GradeVdd);
                    continue;
                }
                if (s.Contains("GRADE", StringComparison.OrdinalIgnoreCase))
                {
                    _ = int.TryParse(s.Split('=').Last(), out searchResultLine.Grade);
                }
            }
            return searchResultLine;
        }

        public static void SetSearchResultbySearchResultWoTestLine(ref SiteInfo[] siteInfoArray, List<SearchResultWoTestLine1> searchResultWoTestLine1s)
        {
            foreach (SearchResultWoTestLine1 searchResultWoTestLine in searchResultWoTestLine1s)
            {
                ResultLine searchResultLine = searchResultWoTestLine.GetSearchResultLine();
                PowerZone powerRef = siteInfoArray[searchResultLine.Site].AllPowers.Find(x => x.PinMode == searchResultLine.Pmode)!;
                powerRef.FinalStep = powerRef.AllSteps.FindIndex(x =>
                            x.EqName == searchResultLine.Eqn && x.Lvcc == searchResultLine.Grade && x.ProductValue == searchResultLine.GradeVdd);
                powerRef.Step = powerRef.FinalStep;
                powerRef.StartStep = powerRef.FinalStep;
                powerRef.StopStep = powerRef.GetPosCount() - 1;
                powerRef.SearchStatus = EnumSearchStatus.Search;
                powerRef.Bin = powerRef.CurrentStep.Bin;
            }
        }
    }

    public class ResultLine : BinCutLineBase
    {
        public int Site;
        public string Pmode = string.Empty;
        public int Eqn;
        public int Grade;
        public int GradeVdd;
    }

}
