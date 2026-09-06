using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlSheets;

using LogLib.Static;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz
{
    public class FlowGenerator
    {
        public HardIpInputData HardIpInputData { get; }

        public FlowGenerator(HardIpInputData hardIpInputData)
        {
            HardIpInputData = hardIpInputData;
        }

        public virtual List<SubFlowSheet> GenFlow(Dictionary<string, HardIpSheet> planDic)
        {
            const string cz2SheetName = "Flow_HARDIP_Char";
            var flowSheets = new ConcurrentBag<SubFlowSheet>();
            var shmooFlowSheets = new ConcurrentBag<SubFlowSheet>();
            var hardIpCharSheet = new SubFlowSheet(cz2SheetName, "HARDIP_CHAR");

            if (LocalSpecs.Options.SplitPatBurstPage)
            {
                var tmpPlanDic = new Dictionary<string, HardIpSheet>();
                foreach (KeyValuePair<string, HardIpSheet> planSheetItems in planDic)
                {
                    var burstItems = planSheetItems.Value.Rows.Where(p => p.Pattern.IsMultiple() || p.Pattern.GetLastPayload().Equals(HardIpConstData.NoPattern, StringComparison.OrdinalIgnoreCase) || (!p.Pattern.IsMultiple() && p.MiscInfo.ContainsIgnoreCase("ref_subblock"))).ToList();
                    var instanceItems = planSheetItems.Value.Rows.Where(p => HardIpConstData.RegInsInPatt.IsMatch(p.Pattern.GetLastPayload())).ToList();
                    var singleItems = planSheetItems.Value.Rows.Where(p => !p.Pattern.IsMultiple() && !HardIpConstData.RegInsInPatt.IsMatch(p.Pattern.GetLastPayload())).ToList();
                    var hardIpSheet = new HardIpSheet
                    {
                        Rows = ReArrangeBurstPatterns(burstItems, instanceItems, singleItems).ToList()
                    };
                    tmpPlanDic.Add(planSheetItems.Key, hardIpSheet);
                    var hardIpSheet1 = new HardIpSheet
                    {
                        Rows = ReArrangeBurstPatterns(singleItems, instanceItems).ToList()
                    };
                    tmpPlanDic.Add(planSheetItems.Key + "_S", hardIpSheet1);
                }
                planDic = tmpPlanDic;
            }

            foreach (string sheetName in planDic.Keys)
            //Parallel.ForEach(planDic.Keys, sheetName =>
            {
                try
                {
                    FlowSheetGeneratorBase flowSheetGenerator = new HardIpFlowSheetGenerator(HardIpInputData, sheetName, planDic[sheetName].Rows);
                    if (SearchInfo.IsHardipIdsSheet(sheetName))
                    {
                        if (LocalSpecs.Options.Device == EnumDevice.AP)
                        {
                            flowSheetGenerator = new IdsFlowSheetGenerator(HardIpInputData, sheetName, planDic[sheetName].Rows);
                        }
                    }

                    if (SearchInfo.IsHardipRtosSheet(sheetName))
                    {
                        flowSheetGenerator = new RtosFlowSheetGenerator(HardIpInputData, sheetName, planDic[sheetName].Rows);
                    }

                    List<SubFlowSheet> subFlowSheets;
                    if (Regex.IsMatch(sheetName, NeededSheets.PrefixRtos, RegexOptions.IgnoreCase))
                    {
                        subFlowSheets = flowSheetGenerator.GenerateFlowSheet();
                    }
                    else
                    {
                        subFlowSheets = CustomWorkflow(ref flowSheetGenerator, planDic, sheetName);
                    }

                    if (subFlowSheets == null)
                    {
                        if (HardIpInputData.HardIpParaData.SplitCzFlow)
                        {
                            subFlowSheets = flowSheetGenerator.GenerateFlowSheetForSplitCz();
                        }
                        else
                        {
                            subFlowSheets = flowSheetGenerator.GenerateFlowSheet();
                        }
                    }

                    foreach (SubFlowSheet subFlowSheet in subFlowSheets)
                    {
                        //Delete jobs if All jobs are enable and replace job name in job column by actual job in config
                        subFlowSheet.FilterFlowJobs(LocalSpecs.AllJobsHardIp);
                        flowSheets.Add(subFlowSheet);
                    }

                    SubFlowSheet subShmooFlowSheet = flowSheetGenerator.GenerateShmooFlowSheet();
                    if (subShmooFlowSheet != null)
                    {
                        shmooFlowSheets.Add(subShmooFlowSheet);
                    }

                    SubFlowSheet vtSubShmooFlowSheet = flowSheetGenerator.GenerateVtShmooFlowSheet();
                    if (vtSubShmooFlowSheet != null)
                    {
                        flowSheets.Add(vtSubShmooFlowSheet);
                    }
                }
                catch (Exception ex)
                {
                    Response.Report("Generating Flow " + sheetName + " failed " + ex.Message, EnumMessageLevel.Error, 0);
                }
            }

            var sheets = flowSheets.OrderBy(x => x.Name).ToList();

            if (shmooFlowSheets.Count > 0)
            {
                foreach (SubFlowSheet sheet in shmooFlowSheets)
                {
                    hardIpCharSheet.AddRows(sheet.Rows);
                }

                hardIpCharSheet.AddRow(FlowRowGeneratorBase.GenReturnRow());
                sheets.Add(hardIpCharSheet);
            }

            return sheets;
        }

        public List<HardIpPattern> ReArrangeBurstPatterns(List<HardIpPattern> patterns, List<HardIpPattern> insertPatterns, List<HardIpPattern> singlePatterns = null)
        {
            var resultList = new List<HardIpPattern>();
            var patternList = patterns.Select((pat, idx) => new { pat, idx }).ToList();
            if (singlePatterns == null && patternList.Count == 0)
            {
                patternList = insertPatterns.Select((pat, idx) => new { pat, idx }).ToList();
            }

            for (int i = 0; i < patternList.Count; i++)
            {
                var currentPat = patternList[i];
                resultList.Add(currentPat.pat);

                if (currentPat.pat.BurstPatterns.Any())
                {
                    var burstList = currentPat.pat.BurstPatterns.Select((pat, idx) => new { pat, idx });
                    foreach (var burstPat in burstList)
                    {
                        if (burstPat.idx <= 0 || singlePatterns == null)
                        {
                            continue;
                        }

                        var singleList = singlePatterns.Select((pat, idx) => new { pat, idx }).ToList();
                        var startPat = singleList.Find(p => p.pat.RowNum == burstPat.pat.RowNum);
                        if (startPat == null)
                        {
                            if (burstPat.pat.RowNum == 0)
                            {
                                var nextPat = patternList.Find(p => p.idx == currentPat.idx + 1);
                                List<HardIpPattern> insertList = insertPatterns.FindAll(p => nextPat == null ? p.RowNum > currentPat.pat.RowNum : p.RowNum > currentPat.pat.RowNum && p.RowNum < nextPat.pat.RowNum);
                                resultList.AddRange(insertList);
                                break;
                            }
                        }
                        else
                        {
                            var endPat = singleList.Find(p => p.idx == startPat.idx + 1);
                            List<HardIpPattern> insertList = insertPatterns.FindAll(p => endPat == null ? p.RowNum > startPat.pat.RowNum : p.RowNum > startPat.pat.RowNum && p.RowNum < endPat.pat.RowNum);
                            resultList.AddRange(insertList);
                        }
                    }
                }
                else
                {
                    var nextPat = patternList.Find(p => p.idx == currentPat.idx + 1);
                    List<HardIpPattern> insertList = insertPatterns.FindAll(p => nextPat == null ? p.RowNum > currentPat.pat.RowNum : p.RowNum > currentPat.pat.RowNum && p.RowNum < nextPat.pat.RowNum);
                    resultList.AddRange(insertList);
                }
            }
            return resultList;
        }

        protected virtual List<SubFlowSheet> CustomWorkflow(ref FlowSheetGeneratorBase flowSheetGenerator, Dictionary<string, HardIpSheet> planDic, string sheetName)
        {
            return null;
        }
    }
}
