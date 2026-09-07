using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow
{
    public class RtosFlowSheetGenerator : HardIpFlowSheetGenerator
    {
        public RtosFlowSheetGenerator(HardIpInputData hardIpInputData, string sheetName, List<HardIpPattern> patternList) : base(hardIpInputData, sheetName, patternList)
        {
            FlowRowGenerator = new RtosFlowRowGenerator(hardIpInputData, SheetName);
        }

        protected override List<FlowRow> GenFlowBodyRows(bool shmooflag = false, bool vtShmooFlag = false)
        {
            var flowBodyRows = new List<FlowRow>();
            List<string> volatages = new List<string>();
            volatages.Add(HardIpConstData.LabelNv);
            volatages.Add(HardIpConstData.LabelLv);
            volatages.Add(HardIpConstData.LabelHv);

            if (vtShmooFlag)
            {
                return flowBodyRows;
            }
            foreach (string votlage in volatages)
            {
                flowBodyRows.AddRange(GenFlowTestRowsByVoltage(votlage));
            }


            return flowBodyRows;
        }


        private static List<FlowRow> SetDeviceCondition(string siteFlag, List<FlowRow> rows)
        {
            string siteVar = siteFlag;
            var ifFlowRows = new List<FlowRow>();

            if (!string.IsNullOrEmpty(siteVar))
            {
                if (siteVar.Contains("&&") || siteVar.Contains("||"))
                {
                    ifFlowRows.Add(FlowRow.GenIfCondition(siteVar, ""));
                    ifFlowRows.AddRange(rows);
                    ifFlowRows.Add(FlowRow.GenEndIf(""));
                }
                else
                {
                    foreach (FlowRow row in rows)
                    {
                        AddCondition(siteVar, row);
                    }

                    ifFlowRows.AddRange(rows);
                }

            }

            return ifFlowRows;
        }

        private static void AddCondition(string siteVar, FlowRow flowRow)
        {
            flowRow.DeviceName = siteVar.TrimStart('!');
            if (!siteVar.Trim().EndsWith("False", StringComparison.CurrentCultureIgnoreCase))
            {
                flowRow.DeviceCondition = siteVar.StartsWith("!") ? "Flag-false" : "Flag-true";
            }
            else
            {
                flowRow.DeviceCondition = siteVar.StartsWith("!") ? "Flag-true" : "Flag-false";
            }
        }
        public override SubFlowSheet GenerateShmooFlowSheet()
        {
            return null;
        }
        protected override List<FlowRow> GenFlowTestRowsByVoltage(string labelVoltage = "", bool shmooCharflag = false, bool vtShmooflag = false)
        {
            var allflowRows = new List<FlowRow>();

            FlowRowGenerator.LabelVoltage = labelVoltage;
            bool isFirstScenario = true;
            string rtosbootInstanceName = "";
            foreach (HardIpPattern pattern in ExtendedPatList)
            {
                try
                {
                    FlowRowGenerator.Pat = pattern;
                    List<FlowRow> flowRows;

                    if (FlowRowGenerator.NoNeedToGen)
                    {
                        continue;
                    }

                    if (pattern.FunctionName.Equals(FuncNameConst.CSharpFuncNameRtosRunScenario, StringComparison.OrdinalIgnoreCase))
                    {
                        flowRows = FlowRowGenerator.GenRunScenarioRows(rtosbootInstanceName, isFirstScenario);
                        if (isFirstScenario)
                        {
                            isFirstScenario = false;
                        }
                        if (!pattern.MiscInfoDict.ContainsKey(HardIpConstData.IgnorePatBinOut))
                        {
                            flowRows.Add(FlowRowGenerator.GenBinTableRow());
                        }

                        allflowRows.AddRange(!string.IsNullOrEmpty(pattern.SiteFlag) ? SetDeviceCondition(pattern.SiteFlag, flowRows) : flowRows);
                        flowRows.Clear();
                    }
                    else
                    {
                        flowRows = FlowRowGenerator.GenTestRows();
                        if (pattern.FunctionName.Equals(FuncNameConst.CSharpFuncNameBootUp, StringComparison.OrdinalIgnoreCase))
                        {
                            if (string.IsNullOrEmpty(rtosbootInstanceName) && flowRows.Any())
                            {
                                rtosbootInstanceName = flowRows.FirstOrDefault().Parameter;
                            }
                        }

                        if (!pattern.MiscInfoDict.ContainsKey(HardIpConstData.IgnorePatBinOut) && !HardIpConstData.RegOpcodeInPatt.IsMatch(pattern.Pattern.RealPatternName))
                        {
                            flowRows.Add(FlowRowGenerator.GenBinTableRow());
                        }
                        allflowRows.AddRange(!string.IsNullOrEmpty(pattern.SiteFlag) ? SetDeviceCondition(pattern.SiteFlag, flowRows) : flowRows);
                        flowRows.Clear();
                    }
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                    throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum);
                }
            }

            FlowRowGenerator.InitOriginalItems.Clear();
            return allflowRows;
        }

    }
}
