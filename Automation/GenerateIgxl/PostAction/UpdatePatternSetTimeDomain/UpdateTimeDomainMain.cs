using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.UpdatePatternSetTimeDomain
{
    public class UpdateTimeDomainMain
    {
        private List<EnumEquipment> _equipmentList;
        public UpdateTimeDomainMain(List<EnumEquipment> equipmentList)
        {
            _equipmentList = equipmentList;
        }

        public void UpdatePatternSetSheets(ref PatSetSheet patSetsAllSheet, ref IEnumerable<PatSetSheet> patSetSheets)
        {
            //Clear .TimeDomain for UltraFlex
            if (_equipmentList.Contains(EnumEquipment.UltraFlex))
            {
                foreach (PatSetRow patSetsAllRow in patSetsAllSheet.Rows.SelectMany(x => x.PatSetRows))
                {
                    patSetsAllRow.TimeDomain = "";
                }
            }
            //Set .TimeDomain for all PatSets sheets for non UltraFlex
            else
            {
                var timeDomainReference = patSetsAllSheet.Rows
                    .GroupBy(x => x.PatSetName)
                    .ToDictionary(
                        sets => sets.Key,
                        sets => string.Join(",", sets.SelectMany(set => set.PatSetRows)
                                                        .Select(x => x.TimeDomain)
                                                        .Distinct(StringComparer.OrdinalIgnoreCase)),
                        StringComparer.OrdinalIgnoreCase);
                foreach (PatSetRow patSetRow in patSetSheets.SelectMany(sheet => sheet.Rows).SelectMany(set => set.PatSetRows))
                {
                    if (timeDomainReference.TryGetValue(patSetRow.File, out string timeDomain))
                    {
                        patSetRow.TimeDomain = timeDomain;
                    }
                }
            }
        }

        public void UpdateTimeSetSheets(ref IEnumerable<TimeSetBasicSheet> timeSetSheets)
        {
            if (_equipmentList.Contains(EnumEquipment.UltraFlex))
            {
                foreach (TimeSetBasicSheet timeSetSheet in timeSetSheets)
                {
                    timeSetSheet.TimeDomain = "";
                }
            }
        }
    }
}
