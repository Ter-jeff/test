using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Utility.Atpg;

namespace Automation.GenerateIgxl.EFuse.Business
{
    public class EfuseFinalInstanceRows : List<EfuseFinalInstanceRow>
    {
        public EfuseFinalInstanceRows ReTestNameDuplicateRows()
        {
            var allGroups = this.Where(x => !string.IsNullOrEmpty(x.TestName)).GroupBy(y => y.TestName).ToList();
            foreach (IGrouping<string, EfuseFinalInstanceRow> group in allGroups)
            {
                for (int i = 0; i < group.Count(); i++)
                {
                    EfuseFinalInstanceRow row1 = group.ElementAt(i);
                    for (int j = i + 1; j < group.Count(); j++)
                    {
                        EfuseFinalInstanceRow row2 = group.ElementAt(j);
                        if (row1.TestName.Equals(row2.TestName, StringComparison.CurrentCultureIgnoreCase))
                        {
                            if (IsInitItem(row1, row2))
                            {
                                if (!AtpgService.IsSamePatternList(row1.EfusePatternRow.InitList, row2.EfusePatternRow.InitList))
                                {
                                    row2.TestNameTemp = row2.TestName + "_RowNum" + row2.EfusePatternRow.RowNum;
                                }
                                else
                                {
                                    // A,B,C => A,C are the same, but B is not. Algen A,C patSetName
                                    row2.TestNameTemp = row1.TestNameTemp;
                                    row2.TestName = row1.TestName;
                                }
                            }
                            else
                            {
                                if (!AtpgService.IsSamePatternList(row1.EfusePatternRow.PayloadList, row2.EfusePatternRow.PayloadList))
                                {
                                    row2.TestNameTemp = row2.TestName + "_RowNum" + row2.EfusePatternRow.RowNum;
                                }
                                else
                                {
                                    // A,B,C => A,C are the same, but B is not. Algen A,C patSetName
                                    row2.TestNameTemp = row1.TestNameTemp;
                                    row2.TestName = row1.TestName;
                                }
                            }
                        }
                    }
                }
            }
            //Set for final PatSetName
            for (int i = 0; i < Count; i++)
            {
                EfuseFinalInstanceRow row = this[i];
                row.TestName = row.TestNameTemp;
            }
            return this;
        }

        public EfuseFinalInstanceRows RePatternNameDuplicateRows()
        {
            var groups = this.Where(x => !string.IsNullOrEmpty(x.PatSet.PatSetName)).GroupBy(y => y.PatSetName).ToList();

            foreach (IGrouping<string, EfuseFinalInstanceRow> group in groups)
            {
                for (int i = 0; i < group.Count(); i++)
                {
                    EfuseFinalInstanceRow row1 = group.ElementAt(i);
                    for (int j = i + 1; j < group.Count(); j++)
                    {
                        EfuseFinalInstanceRow row2 = group.ElementAt(j);
                        if (IsInitItem(row1, row2))
                        {
                            if (!AtpgService.IsSamePatternList(row1.EfusePatternRow.InitList, row2.EfusePatternRow.InitList))
                            {
                                if (row2.PatSetName == row1.PatSetName && !row2.PatSetName.Contains("Init_ALL"))
                                {
                                    row2.PatSetNameTemp = row2.PatSetName + "_RowNum" + row2.EfusePatternRow.RowNum;
                                }
                            }
                            else
                            {
                                // A,B,C => A,C are the same, but B is not. Algen A,C patSetName
                                row2.PatSetNameTemp = row1.PatSetNameTemp;
                                row2.PatSetName = row1.PatSetName;
                            }
                        }
                        else
                        {
                            if (!AtpgService.IsSamePatternList(row1.EfusePatternRow.PayloadList, row2.EfusePatternRow.PayloadList))
                            {
                                if (row2.PatSetName == row1.PatSetName && !row2.PatSetName.Contains("Init_ALL"))
                                {
                                    row2.PatSetNameTemp = row2.PatSetName + "_RowNum" + row2.EfusePatternRow.RowNum;
                                }
                            }
                            else
                            {
                                // A,B,C => A,C are the same, but B is not. Algen A,C patSetName
                                row2.PatSetNameTemp = row1.PatSetNameTemp;
                                row2.PatSetName = row1.PatSetName;
                            }
                        }
                    }
                }
            }

            //Set for final PatSetName
            for (int i = 0; i < Count; i++)
            {
                EfuseFinalInstanceRow row = this[i];
                row.PatSetName = row.PatSetNameTemp;
            }
            return this;
        }

        private bool IsInitItem(EfuseFinalInstanceRow row1, EfuseFinalInstanceRow row2)
        {
            if (!string.IsNullOrEmpty(row1.InitPatName) && !string.IsNullOrEmpty(row2.InitPatName))
            {
                return true;
            }

            return false;
        }
    }
}
