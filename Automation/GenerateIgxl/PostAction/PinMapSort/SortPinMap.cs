using System;
using System.Collections.Generic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.PinMapSort
{
    public class SortPinMap
    {
        private const string ConTimeDomain = "TimeDomain";
        public void Sort(PinMapSheet pPinMapSheet)
        {
            var timeDomainList = new List<PinGroup>();
            for (int i = 0; i < pPinMapSheet.GroupList.Count; i++)
            {
                if (pPinMapSheet.GroupList[i].PinType.Equals(ConTimeDomain, StringComparison.OrdinalIgnoreCase))
                {
                    timeDomainList.Add(pPinMapSheet.GroupList[i]);
                    pPinMapSheet.RemoveGroupAt(i);
                    i--;
                }
            }
            pPinMapSheet.AddGroups(timeDomainList);
        }
    }
}
