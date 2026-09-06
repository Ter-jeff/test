using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using ProjectConfigLib.ProjectConfig;

using TestPlanLib.Static;

namespace RfLib.Dvdc.GenFlow
{
    public partial class DvdcFlowGenerator(HardIpInputData hardIpInputData) : FlowGenerator(hardIpInputData)
    {
        [GeneratedRegex(NeededSheets.PrefixWireless, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("wireless_|lcd_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(NeededSheets.PrefixRtos, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(NeededSheets.PrefixLcd, RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex("wireless_|lcd_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex4();

        public override List<SubFlowSheet> GenFlow(Dictionary<string, HardIpSheet> planDic)
        {
            const string cz2SheetName = "Flow_HARDIP_Char";
            var flowSheets = new List<SubFlowSheet>();
            var shmooflowSheets = new List<SubFlowSheet>();
            var hardIpCharSheet = new SubFlowSheet(cz2SheetName, "HARDIP_CHAR");

            if (ProjectConfigSingleton.Instance()
                .GetValue("HardIP", "SplitPatBurstPage")
.EqualsIgnoreCase("TRUE"))
            {
                var tmpPlanDic = new Dictionary<string, HardIpSheet>();
                foreach (KeyValuePair<string, HardIpSheet> planSheetItems in planDic)
                {
                    var burstItems =
                        planSheetItems.Value.Rows.Where(
                            p => p.Pattern.IsMultiple() ||
                                 p.Pattern.GetLastPayload().EqualsIgnoreCase(HardIpConstData.NoPattern) ||
                                 (!p.Pattern.IsMultiple() && p.MiscInfo.ContainsIgnoreCase("ref_subblock"))).ToList();
                    var instanceItems = planSheetItems.Value.Rows.Where(
                        p => HardIpConstData.RegInsInPatt.IsMatch(p.Pattern.GetLastPayload())).ToList();

                    var singleItems = planSheetItems.Value.Rows.Where(p => !p.Pattern.IsMultiple() && !HardIpConstData.RegInsInPatt.IsMatch(p.Pattern.GetLastPayload())).ToList();
                    //BurstItems.ForEach(p=>p.SheetName = p.SheetName + "_MG");
                    var hardIpSheet = new HardIpSheet
                    {
                        Rows = [.. ReArrangeBurstPatterns(burstItems, instanceItems, singleItems)]
                    };
                    tmpPlanDic.Add(planSheetItems.Key, hardIpSheet);
                    var hardIpSheet1 = new HardIpSheet
                    {
                        Rows = [.. ReArrangeBurstPatterns(singleItems, instanceItems)]
                    };
                    tmpPlanDic.Add(planSheetItems.Key + "_S", hardIpSheet1);
                }
                planDic = tmpPlanDic;
            }

            foreach (string sheetName in planDic.Keys)
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

                    var subFlowSheets = new List<SubFlowSheet>();
                    if (MyRegex().IsMatch(sheetName))
                    {
                        var subFlowRows = new List<FlowRow>();
                        foreach (string lable in HardIpConstData.LabelVolList)
                        {
                            flowSheetGenerator = new WirelessFlowSheetGenerator(HardIpInputData, MyRegex1().Replace(sheetName, ""), lable, planDic[sheetName].Rows);
                            subFlowSheets.AddRange(flowSheetGenerator.GenerateFlowSheet());
                            subFlowRows.Add(new FlowRow() { Opcode = OpCode.Call, Enable = "HardIP_" + lable, Parameter = flowSheetGenerator.FlowSheetName });
                        }

                        var rFflowSheetGenerator = new WirelessFlowSheetGenerator(HardIpInputData, MyRegex1().Replace(sheetName, ""), "", null!);
                        subFlowSheets.AddRange(rFflowSheetGenerator.GenerateFlowSheetForRfMain(subFlowRows));
                    }
                    else if (MyRegex2().IsMatch(sheetName))
                    {
                        subFlowSheets = flowSheetGenerator.GenerateFlowSheet();
                    }
                    else
                    {
                        if (MyRegex3().IsMatch(sheetName))
                        {
                            flowSheetGenerator = new WirelessFlowSheetGenerator(HardIpInputData, MyRegex4().Replace(sheetName, ""), "", planDic[sheetName].Rows);
                        }

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
                        shmooflowSheets.Add(subShmooFlowSheet);
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

            if (shmooflowSheets.Count > 0)
            {
                foreach (SubFlowSheet sheet in shmooflowSheets)
                {
                    hardIpCharSheet.AddRows(sheet.Rows);
                }

                hardIpCharSheet.AddRow(FlowRowGeneratorBase.GenReturnRow());
                flowSheets.Add(hardIpCharSheet);
            }

            return flowSheets;
        }
    }
}
