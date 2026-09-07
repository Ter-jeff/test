using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.Singleton;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Utility;

namespace Automation.GenerateIgxl.BinCut.Business
{
    internal class BinCutPatSetWriter
    {
        public List<PatSetSheet> GetPatSetSheet(List<BinCutFinalInstanceRow> bincutInstanceRows, HashSet<string> regularPatsetRows = null)
        {
            List<PatSetSheet> patSetSheets = GenPatternSetSheet(bincutInstanceRows, regularPatsetRows);

            patSetSheets = AddCommmadandFlagInPatSet(patSetSheets, bincutInstanceRows);

            return patSetSheets;
        }

        private List<PatSetSheet> GenPatternSetSheet(List<BinCutFinalInstanceRow> bincutInstanceRows, HashSet<string> regularPatsetRows = null)
        {
            string sheetName = "PatSets_BinCut";
            var patSetSheets = new List<PatSetSheet>();
            if (bincutInstanceRows.First().BinCutInstanceRow.SheetName.ContainsIgnoreCase("post"))
            {
                sheetName += "_OutsideBV";

                bincutInstanceRows = bincutInstanceRows
                    .Where(r =>
                        (string.IsNullOrWhiteSpace(r.InitPatSetName) || !regularPatsetRows.Contains(r.InitPatSetName.Trim())) &&
                        (string.IsNullOrWhiteSpace(r.PatSetName) || !regularPatsetRows.Contains(r.PatSetName.Trim()))
                    )
                    .ToList();
            }

            var patSetSheet = new PatSetSheet(sheetName);
            var existPatSetList = new List<string>();
            int i = 0;
            int patsetRowCnt = 0;
            foreach (BinCutFinalInstanceRow bincutInstance in bincutInstanceRows)
            {
                if (bincutInstance.BinCutInstanceRow.Type == BincutInstanceType.Hardip || bincutInstance.BinCutInstanceRow.Type == BincutInstanceType.Rtos)
                {
                    continue;
                }

                if (patsetRowCnt > 90000)
                {
                    patsetRowCnt = 0;
                    i++;
                    patSetSheets.Add(patSetSheet);
                    patSetSheet = new PatSetSheet(sheetName + "_" + i);
                }
                string comment = bincutInstance.BinCutInstanceRow.FlowName + " , " + bincutInstance.BinCutInstanceRow.SheetName + ": RowNum" + bincutInstance.BinCutInstanceRow.RowNum;

                //init pattern set
                if (BinCutInstanceRowUtility.IsBist(bincutInstance.BinCutInstanceRow.FlowName))
                {
                    if (bincutInstance.InitPatSet != null && !existPatSetList.Exists(x => x.Equals(bincutInstance.InitPatSet.PatSetName)))
                    {
                        if (bincutInstance.InitList.Count != 1)
                        {
                            patSetSheet.AddRow(bincutInstance.InitPatSet);
                            patsetRowCnt += bincutInstance.InitPatSet.PatSetRows.Count;
                            existPatSetList.Add(bincutInstance.InitPatSet.PatSetName);
                        }
                    }
                }

                //PL pattern set
                if (bincutInstance.PatSet != null && !existPatSetList.Exists(x => x.Equals(bincutInstance.PatSet.PatSetName)))
                {
                    if (bincutInstance.PatSet.PatSetRows.Count > 0) //Only one playload => don't need create patSet for bist
                    {
                        bincutInstance.PatSet.PatSetRows.First().Comment = comment;
                        patSetSheet.AddRow(bincutInstance.PatSet);
                        patsetRowCnt += bincutInstance.PatSet.PatSetRows.Count;
                        existPatSetList.Add(bincutInstance.PatSet.PatSetName);
                    }
                }
            }
            patSetSheets.Add(patSetSheet);
            return patSetSheets;
        }

        private List<PatSetSheet> AddCommmadandFlagInPatSet(List<PatSetSheet> patSetSheets, List<BinCutFinalInstanceRow> bincutInstanceRows)
        {
            //patSetSheet: Results of all items in the bincut instance sheet
            //bincutInstanceRows: The result is that the flow sheet conditionally exists in the BinCut instance sheet
            var patSetHashSet = new HashSet<string>(bincutInstanceRows.Where(x => x.IsUsed && x.PatSet != null && !string.IsNullOrEmpty(x.PatSet.PatSetName)).Select(x => x.PatSet.PatSetName), StringComparer.CurrentCultureIgnoreCase);
            var initPatSetHashSet = new HashSet<string>(bincutInstanceRows.Where(x => x.IsUsed && x.InitPatSet != null && !string.IsNullOrEmpty(x.InitPatSet.PatSetName)).Select(x => x.InitPatSet.PatSetName), StringComparer.CurrentCultureIgnoreCase);
            foreach (string initPatSet in initPatSetHashSet)
            {
                patSetHashSet.Add(initPatSet);
            }

            foreach (PatSetSheet patSetSheet in patSetSheets)
            {
                if (patSetSheet.Name.StartsWith("PatSets_BinCut", StringComparison.OrdinalIgnoreCase))
                {
                    for (int i = 0; i < patSetSheet.Rows.Count; i++)
                    {
                        if (!patSetHashSet.Contains(patSetSheet.Rows[i].PatSetName))
                        {
                            patSetSheet.Rows[i].IsBackup = true;
                            foreach (PatSetRow row in patSetSheet.Rows[i].PatSetRows)
                            {
                                row.Comment += ", dont_useInFlow";
                                row.IsBackup = true;
                            }
                        }

                        foreach (PatSetRow item in patSetSheet.Rows[i].PatSetRows)
                        {
                            Dictionary<string, PatternData> patternDatas = AcTSetCategoryMapSingleton.Instance().PatternList;
                            if (string.IsNullOrEmpty(item.File))
                            {
                                continue;
                            }

                            string patternName = item.File.ToLower();
                            if (patternDatas.ContainsKey(patternName))
                            {
                                if (patternDatas[patternName].Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
                                {
                                    item.Comment += ", dont_useInCsv";
                                    item.IsBackup = true;
                                }
                                else
                                {
                                    if (patternDatas[patternName].FileVersion.Equals("n/a", StringComparison.OrdinalIgnoreCase))
                                    {
                                        item.Comment += ", no_pattern";
                                        item.IsBackup = true;
                                    }
                                    else
                                    {
                                        if (!patternDatas[patternName].IsExist)
                                        {
                                            item.Comment += ", no_pattern";
                                            item.IsBackup = true;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                item.Comment += ", no_pattern";
                                item.IsBackup = true;
                            }
                        }
                    }
                }
            }

            return patSetSheets;
        }
    }
}
