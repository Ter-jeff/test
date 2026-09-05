using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.PatSetProcessing
{
    public class SplitPatSet
    {
        public readonly int MaxItemsPerContainer;


        public SplitPatSet(int maxItemsPerContainer = 4000)
        {
            if (maxItemsPerContainer <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(maxItemsPerContainer));
            }
            MaxItemsPerContainer = maxItemsPerContainer;
        }



        public void SplitPatset(Dictionary<string, PatSetSheet> patSetSheets)
        {

            if (patSetSheets == null || patSetSheets.Count == 0)
            {
                return;
            }

            var splitPatSetSheets = new List<string>();
            var multiplePatSetSheets = new Dictionary<string, List<PatSetSheet>>(StringComparer.OrdinalIgnoreCase);

            foreach (KeyValuePair<string, PatSetSheet> a in patSetSheets)
            {

                string path_patSet = Path.GetDirectoryName(a.Key);


                var mainList = a.Value.Rows.Where(x => x.PatSetRows.All(y => !y.IsBackup)).ToList();
                var backupList = a.Value.Rows.Where(x => x.PatSetRows.Any(y => y.IsBackup)).ToList();

                int totalRowCount = a.Value.Rows.Sum(p => p.PatSetRows.Count);

                if (totalRowCount <= MaxItemsPerContainer)
                {
                    continue;
                }

                if (mainList.Count == 0 && backupList.Count > 0)
                {
                    continue;
                }

                int sheetIdx = 1;
                int patSetCount = 0;


                if (!multiplePatSetSheets.TryGetValue(path_patSet, out List<PatSetSheet> list))
                {
                    list = new List<PatSetSheet>();
                    multiplePatSetSheets.Add(path_patSet, list);
                }


                var newSheet = new PatSetSheet($"{a.Value.Name}_{sheetIdx}");
                list.Add(newSheet);


                foreach (PatSet patSet in mainList)
                {
                    if ((patSet.PatSetRows.Count() + patSetCount) > MaxItemsPerContainer)
                    {
                        sheetIdx++;
                        patSetCount = 0;

                        newSheet = new PatSetSheet($"{a.Value.Name}_{sheetIdx}");
                        list.Add(newSheet);
                    }

                    patSetCount += patSet.PatSetRows.Count();
                    newSheet.Rows.Add(patSet);
                }

                foreach (PatSet patSet in backupList)
                {
                    if ((patSet.PatSetRows.Count() + patSetCount) > MaxItemsPerContainer)
                    {
                        sheetIdx++;
                        patSetCount = 0;

                        newSheet = new PatSetSheet($"{a.Value.Name}_{sheetIdx}");
                        list.Add(newSheet);
                    }

                    patSetCount += patSet.PatSetRows.Count();
                    newSheet.Rows.Add(patSet);
                }

                splitPatSetSheets.Add(Path.Combine(path_patSet, a.Value.Name));

            }

            foreach (string b in splitPatSetSheets)
            {
                patSetSheets.Remove(b);
            }


            foreach (KeyValuePair<string, List<PatSetSheet>> kvp in multiplePatSetSheets)
            {
                string key = kvp.Key;
                List<PatSetSheet> sheets = kvp.Value;

                foreach (PatSetSheet sheet in sheets)
                {
                    patSetSheets.Add(Path.Combine(key, sheet.Name), sheet);
                }
            }

        }

    }
}
