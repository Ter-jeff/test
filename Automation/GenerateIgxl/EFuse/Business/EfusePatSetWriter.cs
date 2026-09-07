using System.Collections.Generic;
using System.Linq;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.EFuse.Business
{
    internal class EfusePatSetWriter
    {
        public PatSetSheet GetPatSetSheet(List<EfuseFinalInstanceRow> efuseInstanceRows)
        {
            return GenPatternSetSheet(efuseInstanceRows);
        }

        private PatSetSheet GenPatternSetSheet(List<EfuseFinalInstanceRow> efuseInstanceRows)
        {
            string sheetName = "PatSets_eFuse";
            var patSetSheet = new PatSetSheet(sheetName);
            foreach (EfuseFinalInstanceRow efuseInstance in efuseInstanceRows)
            {
                if (efuseInstance.EfusePatternRow == null)
                {
                    continue;
                }

                List<PatSet> matchedPatSets = new List<PatSet>();
                if (!string.IsNullOrEmpty(efuseInstance.PatSet.PatSetName))
                {
                    string efuseName = efuseInstance.PatSet.PatSetName.Trim();
                    if (patSetSheet.Rows.Any(c => efuseName.Equals(c.PatSetName)))
                    {
                        matchedPatSets.Add(efuseInstance.PatSet);
                    }
                }
                else
                {
                    continue;
                }

                bool isDuplicate = false;
                foreach (PatSet matchedPatSet in matchedPatSets)
                {
                    if (patSetSheet.IsExistTheSamePatSet(matchedPatSet, out _))
                    {
                        isDuplicate = true;
                        break;
                    }
                }

                if (!isDuplicate)
                {
                    string comment = efuseInstance.EfusePatternRow.SheetName + ": RowNum" + efuseInstance.EfusePatternRow.RowNum;
                    if (matchedPatSets.Any())
                    {
                        efuseInstance.PatSet.PatSetName += "_RowNum" + efuseInstance.EfusePatternRow.RowNum;
                    }
                    //PL pattern set
                    if (efuseInstance.PatSet != null && !string.IsNullOrEmpty(efuseInstance.PatSet.PatSetName) && !patSetSheet.IsExist(efuseInstance.PatSet.PatSetName))
                    {
                        if (efuseInstance.PatSet.PatSetRows.Count > 0 && !string.IsNullOrEmpty(efuseInstance.PatSet.PatSetName))
                        {
                            efuseInstance.PatSet.PatSetRows.First().Comment = comment;
                            patSetSheet.AddRow(efuseInstance.PatSet);
                        }
                    }
                }
            }
            return patSetSheet;
        }
    }
}
