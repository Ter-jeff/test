using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Utility.Atpg;

using CommonLib.Utility;

namespace Automation.GenerateIgxl.BinCut.Base
{
    public class BinCutFinalInstanceRows : List<BinCutFinalInstanceRow>
    {
        public BinCutFinalInstanceRows RePatSetNameDuplicateRows()
        {
            //Rename when length > 255
            for (int i = 0; i < Count; i++)
            {
                BinCutFinalInstanceRow row = this[i];
                if (row.PatSetName.Length > 255 && !string.IsNullOrEmpty(row.BinCutInstanceRow.Instance))
                {
                    string name = Combination.CombineByUnderLine(new List<string> { row.Domain + row.Block, "Inst", row.BinCutInstanceRow.Instance });
                    row.PatSetName = name;
                    row.PatSetNameTemp = name;
                }
            }

            var groups = this.GroupBy(x => x.PatSetName.ToUpper()).ToList();
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                if (group.Count() == 1)
                {
                    continue;
                }

                if (group.Any(x => !string.IsNullOrEmpty(x.BinCutInstanceRow.PatSetNameOrange)))
                {
                    continue;
                }

                GetUniqueNameByPayload(group);
            }

            groups = this.GroupBy(x => x.PatSetName.ToUpper()).ToList();
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                if (group.Count() == 1)
                {
                    continue;
                }

                if (group.Any(x => !string.IsNullOrEmpty(x.BinCutInstanceRow.PatSetNameOrange)))
                {
                    continue;
                }

                GetUniqueNameByInit(group);
            }

            groups = this.GroupBy(x => x.PatSetName.ToUpper()).ToList();
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                for (int i = 0; i < group.Count(); i++)
                {
                    BinCutFinalInstanceRow row1 = group.ElementAt(i);
                    for (int j = i + 1; j < group.Count(); j++)
                    {
                        BinCutFinalInstanceRow row2 = group.ElementAt(j);
                        if (!AtpgService.IsSamePatternList(row1.PatternList, row2.PatternList))
                        {
                            if (row2.PatSetNameTemp.Equals(row1.PatSetName, StringComparison.CurrentCultureIgnoreCase))
                            {
                                row2.PatSetNameTemp = row2.PatSetName + "_RowNum" + row2.BinCutInstanceRow.RowNum;
                            }
                        }
                        else
                        {
                            // A,B,C => A,C are the same, but B is not. Align A,C patSetName
                            row2.PatSetNameTemp = row1.PatSetNameTemp;
                            row2.PatSetName = row1.PatSetName;
                        }
                    }
                }
            }

            //Set for final PatSetName
            for (int i = 0; i < Count; i++)
            {
                BinCutFinalInstanceRow row = this[i];
                row.PatSetName = row.PatSetNameTemp;
            }
            return this;
        }

        public BinCutFinalInstanceRows RePatSetNameDuplicateRowsBySerialNum()
        {
            var groups = this.GroupBy(x => x.FinalInstName.ToUpper()).ToList();
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                int serial = 0;
                if (group.Count() == 1)
                {
                    continue;
                }

                for (int i = 1; i < group.Count(); i++)
                {
                    serial++;
                    BinCutFinalInstanceRow row = group.ElementAt(i);
                    row.FinalInstName = Combination.CombineByUnderLine(row.FinalInstName, serial.ToString());
                }
            }
            return this;
        }

        public BinCutFinalInstanceRows RePatSetNameDuplicateRowsByKey(List<BinCutFinalInstanceRow> binCutInstanceRows, string key)
        {
            if (binCutInstanceRows.Count == 0)
            {
                return this;
            }

            for (int i = 0; i < Count; i++)
            {
                BinCutFinalInstanceRow postRow = this[i];
                if (binCutInstanceRows.Exists(x => x.PatSetNameTemp.Equals(postRow.PatSetNameTemp, StringComparison.OrdinalIgnoreCase)))
                {
                    postRow.PatSetNameTemp = postRow.PatSetNameTemp + "_" + key;
                    postRow.PatSetName = postRow.PatSetNameTemp;
                    if (postRow.InitPatSet != null)
                    {
                        if (!string.IsNullOrEmpty(postRow.InitPatSet.PatSetName))
                        {
                            postRow.InitPatSet.PatSetName = postRow.InitPatSet.PatSetName + "_" + key;
                        }
                    }
                }
            }
            return this;
        }

        private void GetUniqueNameByPayload(IGrouping<string, BinCutFinalInstanceRow> groups)
        {
            var list = groups.SelectMany(x => x.PayloadList).Distinct().ToList();
            if (!list.Any())
            {
                list = groups.SelectMany(x => x.PatternList).Distinct().ToList();
            }

            if (!list.Any())
            {
                return;
            }

            int index = GetOneKeyIndex(list);
            if (index != -1)
            {
                foreach (BinCutFinalInstanceRow row in groups)
                {
                    row.PayloadName = row.PayloadList.Any() ? GetUniqueName(row.PayloadList, index) : GetUniqueName(row.PatternList, index);
                    row.PatSetName = Combination.CombineByUnderLine(new List<string> { row.Domain + row.Block, row.ModeByFlowName, row.InitName, row.PayloadName });
                    row.PatSetNameTemp = row.PatSetName;
                }
            }
        }

        private void GetUniqueNameByInit(IGrouping<string, BinCutFinalInstanceRow> groups)
        {
            GetExtraInitIndex(groups, out int initIndex, out int segmentIndex);
            if (initIndex != -1 && segmentIndex != -1)
            {
                foreach (BinCutFinalInstanceRow row in groups)
                {
                    string[] arr = row.InitList.ElementAt(initIndex).Split('_');
                    string uniqueInit = arr.Length > segmentIndex ? arr.ElementAt(segmentIndex) : "";
                    row.InitName = Combination.CombineByUnderLine(row.InitName, uniqueInit);
                    row.PatSetName = Combination.CombineByUnderLine(new List<string> { row.Domain + row.Block, row.ModeByFlowName, row.InitName, row.PayloadName });
                    row.PatSetNameTemp = row.PatSetName;
                }
            }
        }

        private void GetExtraInitIndex(IGrouping<string, BinCutFinalInstanceRow> groups, out int initIndex, out int segmentIndex)
        {
            initIndex = -1;
            segmentIndex = -1;
            var counts = groups.Select(x => x.InitList.Count).Distinct().ToList();
            if (counts.Count == 1) //Init count are the same
            {
                int diffInitCount = 0;
                for (int i = 0; i < counts.First(); i++)
                {
                    var list = groups.Select(x => x.InitList.ElementAt(i)).Distinct().ToList();
                    if (list.Count != 1) //All the same
                    {
                        diffInitCount++;
                    }
                }

                if (diffInitCount == 1)
                {
                    for (int i = 0; i < counts.First(); i++)
                    {
                        var list = groups.Select(x => x.InitList.ElementAt(i)).Distinct().ToList();
                        if (list.Count != 1) //All the same
                        {
                            int index = GetOneKeyIndex(list);
                            if (index != -1)
                            {
                                initIndex = i;
                                segmentIndex = index;
                                return;
                            }
                        }
                    }
                }
            }
        }

        private int GetOneKeyIndex(List<string> list)
        {
            var ints = new List<int>();
            var items = list.Select(x => x.Split('_')).ToList();
            int max = items.Max(x => x.Length);
            for (int i = 0; i < max; i++)
            {
                var segments = new List<string>();
                foreach (string[] item in items)
                {
                    if (item.Length > i)
                    {
                        segments.Add(item.ElementAt(i));
                    }
                    else
                    {
                        segments.Add("");
                    }
                }
                segments = segments.Distinct().ToList();
                if (segments.Count != 1)
                {
                    ints.Add(i);
                    break;
                }
            }
            if (ints.Count == 1)
            {
                return ints.First();
            }

            return -1;
        }

        private string GetUniqueName(List<string> payloadList, int index)
        {
            var segments = new List<string>();
            foreach (string payload in payloadList)
            {
                string[] array = payload.Split('_');
                if (array.Length > index)
                {
                    segments.Add(array.ElementAt(index));
                }
            }

            List<string> arr = payloadList.First().Split('_').ToList();
            if (segments.Count >= 1)
            {
                if (arr.Count <= index + 1)
                {
                    arr.AddRange(segments.Distinct());
                }
                else
                {
                    segments.RemoveAt(0);
                    arr.Insert(index + 1, string.Join("_", segments.Distinct()));
                }
            }
            return string.Join("_", arr);
        }

        public BinCutFinalInstanceRows RePatSetNameDuplicateRowBinCut(List<BinCutFinalInstanceRow> previousRows = null)
        {
            List<BinCutFinalInstanceRow> allRows = previousRows != null
                    ? previousRows.Concat(this).ToList()
                    : this.ToList();

            //Rename when length > 255
            for (int i = 0; i < allRows.Count; i++)
            {
                BinCutFinalInstanceRow row = allRows[i];
                if (row.PatSetName.Length > 255 && !string.IsNullOrEmpty(row.BinCutInstanceRow.Instance))
                {
                    string name = Combination.CombineByUnderLine(new List<string> { row.Domain + row.Block, "Inst", row.BinCutInstanceRow.Instance });
                    row.PatSetName = name;
                    row.PatSetNameTemp = name;
                }
            }

            // PayloadList and PatternList
            var patDict = new Dictionary<string, List<string>>(StringComparer.CurrentCultureIgnoreCase);
            var groups = allRows.GroupBy(x => x.PatSetName.ToUpper()).ToList();

            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                foreach (BinCutFinalInstanceRow row in group)
                {
                    string basePatName = row.PatSetName;
                    bool matched = false;

                    List<string> currentList = row.Block == "Mbist" ? row.PayloadList : row.PatternList;

                    foreach (KeyValuePair<string, List<string>> kvp in patDict)
                    {
                        if (AtpgService.IsSamePatternList(kvp.Value, currentList))
                        {
                            row.PatSetName = kvp.Key;
                            row.PatSetNameTemp = kvp.Key;
                            matched = true;
                            break;
                        }
                    }

                    if (!matched)
                    {
                        string newPatName = basePatName;
                        int suffix = 1;

                        while (patDict.ContainsKey(newPatName))
                        {
                            if (AtpgService.IsSamePatternList(patDict[newPatName], currentList))
                            {
                                row.PatSetName = newPatName;
                                row.PatSetNameTemp = newPatName;
                                matched = true;
                                break;
                            }

                            newPatName = basePatName + "_" + suffix;
                            suffix++;
                            if (previousRows != null)
                            {
                                newPatName += "_PBC";
                            }
                        }

                        if (!matched)
                        {
                            row.PatSetName = newPatName;
                            row.PatSetNameTemp = newPatName;
                            patDict[newPatName] = currentList;
                        }
                    }
                }
            }
            var dict = new Dictionary<string, List<string>>(StringComparer.CurrentCultureIgnoreCase);
            var initGroups = allRows.GroupBy(x => x.InitPatSetNameNew.ToUpper()).ToList();
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in initGroups)
            {
                foreach (BinCutFinalInstanceRow row in group)
                {
                    string basePatName = row.InitPatSetNameNew;
                    bool matched = false;

                    if (!string.IsNullOrEmpty(basePatName))
                    {
                        foreach (KeyValuePair<string, List<string>> kvp in patDict)
                        {
                            if (AtpgService.IsSamePatternList(kvp.Value, row.InitList))
                            {
                                row.InitPatSetNameNew = kvp.Key;
                                row.InitPatSetNameNewTemp = kvp.Key;
                                matched = true;
                                break;
                            }
                        }

                        if (!matched)
                        {
                            string newPatName = basePatName;
                            int suffix = 1;

                            while (dict.ContainsKey(newPatName))
                            {
                                if (AtpgService.IsSamePatternList(dict[newPatName], row.InitList))
                                {
                                    row.InitPatSetNameNew = newPatName;
                                    row.InitPatSetNameNewTemp = newPatName;
                                    matched = true;
                                    break;
                                }

                                newPatName = basePatName + "_" + suffix;
                                suffix++;
                                if (previousRows != null)
                                {
                                    newPatName += "_PBC";
                                }
                            }

                            if (!matched)
                            {
                                row.InitPatSetNameNew = newPatName;
                                row.InitPatSetNameNewTemp = newPatName;
                                dict[newPatName] = row.InitList;
                            }
                        }
                    }
                }
            }
            return this;
        }
    }
}
