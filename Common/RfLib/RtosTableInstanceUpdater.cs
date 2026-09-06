using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Reader.ConfigFile.RtosTable;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace RfLib
{
    internal static partial class RtosTableInstanceUpdater
    {
        [GeneratedRegex("^IDS_")]
        private static partial Regex MyRegex();

        internal static void UpdateRtosTableInfoInInstanceSheets()
        {
            ExcelWorksheet? rtosTableSheet = EpWorkbook.TestPlanWorkbook.Worksheets.ToList().Find(x => x.Name.EqualsIgnoreCase("Rtos_Table"));
            if (rtosTableSheet == null)
            {
                return;
            }

            var rtosTableInfos = RtosTableSheet.LoadConfig(rtosTableSheet);

            ApplyRtosTableArgsToInstanceRows(rtosTableInfos);
            InsertRtosIdsRows(rtosTableInfos);
        }

        private static void ApplyRtosTableArgsToInstanceRows(RtosTableSheet rtosTableSheet)
        {
            foreach (IGrouping<string, RtosTableArgRow> sheetnamegrp in rtosTableSheet.ArgRows.GroupBy(x => x.SheetName))
            {
                //Get instance sheet
                InstanceSheet? instSheet = TestProgram.IgxlWorkBk.InsSheets.Values.ToList()
                    .Find(p => p.Name.EqualsIgnoreCase(sheetnamegrp.Key));

                if (instSheet == null)
                {
                    continue;
                }

                foreach (IGrouping<string, RtosTableArgRow> funcnamegrp in sheetnamegrp.GroupBy(x => x.FuncName))
                {
                    foreach (InstanceRow insrow in instSheet.Rows)
                    {
                        //Get function name
                        if (!insrow.VbtName.EqualsIgnoreCase(funcnamegrp.Key))
                        {
                            continue;
                        }

                        List<string> insArgs = [.. insrow.ArgList.Split(',')];
                        foreach (RtosTableArgRow arg in funcnamegrp)
                        {
                            ApplyRtosArgToInstanceRow(insrow, insArgs, arg);
                        }
                    }
                }
            }
        }

        private static void ApplyRtosArgToInstanceRow(InstanceRow instanceRow, List<string> insArgs, RtosTableArgRow rtosTableArgRow)
        {
            if (rtosTableArgRow.ArgName == "DC Category" || rtosTableArgRow.ArgName == "DCCategory")
            {
                instanceRow.DcCategory = rtosTableArgRow.ArgInfo;
            }
            else if (rtosTableArgRow.ArgName == "DC Selector" || rtosTableArgRow.ArgName == "DCSelector")
            {
                instanceRow.DcSelector = rtosTableArgRow.ArgInfo;
            }
            else if (rtosTableArgRow.ArgName == "AC Category" || rtosTableArgRow.ArgName == "ACCategory")
            {
                instanceRow.AcCategory = rtosTableArgRow.ArgInfo;
            }
            else if (rtosTableArgRow.ArgName == "AC Selector" || rtosTableArgRow.ArgName == "ACSelector")
            {
                instanceRow.AcSelector = rtosTableArgRow.ArgInfo;
            }
            else if (rtosTableArgRow.ArgName == "Time Sets" || rtosTableArgRow.ArgName == "TimeSets")
            {
                instanceRow.TimeSets = rtosTableArgRow.ArgInfo;
            }
            else if (rtosTableArgRow.ArgName == "Pin Levels" || rtosTableArgRow.ArgName == "PinLevels")
            {
                instanceRow.PinLevels = rtosTableArgRow.ArgInfo;
            }
            else
            {
                int seqIndex = insArgs.FindIndex(s => s.EqualsIgnoreCase(rtosTableArgRow.ArgName));
                if (seqIndex != -1)
                {
                    instanceRow.Args[seqIndex] = rtosTableArgRow.ArgInfo;
                }
            }
        }

        private static void InsertRtosIdsRows(RtosTableSheet rtosTableSheet)
        {
            foreach (KeyValuePair<string, string> measrow in rtosTableSheet.MeasRows)
            {
                FlowRows flowIdSrows = CollectRtosIdsFlowRows(measrow.Value);
                InstanceRows instIdSrows = CollectRtosIdsInstRows(measrow.Value);

                InsertRtosIdsFlowRows(measrow.Key, flowIdSrows);
                InsertRtosIdsInstRows(measrow.Key, instIdSrows);
            }
        }

        private static FlowRows CollectRtosIdsFlowRows(string measValue)
        {
            var flowIdSrows = new FlowRows();
            foreach (KeyValuePair<string, SubFlowSheet> flowsheet in TestProgram.IgxlWorkBk.SubFlowSheets)
            {
                if (flowsheet.Key.Contains("DCTEST_IDS", StringComparison.CurrentCulture))
                {
                    List<FlowRow> measIdsFlowRows = [.. flowsheet.Value.Rows.FindAll(flowrow =>
                            flowrow.Parameter.EqualsIgnoreCase(measValue))
                        .Select(flowrow => flowrow.Copy())];
                    if (measIdsFlowRows.Count != 0)
                    {
                        foreach (FlowRow row in measIdsFlowRows)
                        {
                            row.Parameter = MyRegex().Replace(row.Parameter, "Rtos_");
                        }

                        flowIdSrows.AddRange(measIdsFlowRows);
                    }
                }
            }

            return flowIdSrows;
        }

        private static InstanceRows CollectRtosIdsInstRows(string measValue)
        {
            var instIdSrows = new InstanceRows();
            foreach (KeyValuePair<string, InstanceSheet> instsheet in TestProgram.IgxlWorkBk.InsSheets)
            {
                if (instsheet.Key.Contains("DCTEST_IDS", StringComparison.CurrentCulture))
                {
                    List<InstanceRow> measIdsInstRows = [.. instsheet.Value.Rows.FindAll(
                            instrow => instrow.TestName.EqualsIgnoreCase(measValue))
                        .Select(instrow => instrow.Copy())];
                    if (measIdsInstRows.Count != 0)
                    {
                        foreach (InstanceRow row in measIdsInstRows)
                        {
                            row.TestName = MyRegex().Replace(row.TestName, "Rtos_");
                            int patidx = row.ArgList.Split(',').ToList().FindIndex(x => x == "patt" || x == "pattern");
                            if (patidx != -1)
                            {
                                row.Args[patidx] = "";
                            }
                        }
                        instIdSrows.AddRange(measIdsInstRows);
                    }
                }
            }

            return instIdSrows;
        }

        private static void InsertRtosIdsFlowRows(string measKey, FlowRows flowRows)
        {
            if (!flowRows.Any())
            {
                return;
            }

            foreach (KeyValuePair<string, SubFlowSheet> flowsheet in TestProgram.IgxlWorkBk.SubFlowSheets)
            {
                if (flowsheet.Key.Contains("Flow_Rtos_UART", StringComparison.CurrentCulture) ||
                    flowsheet.Key.Contains("Flow_Rtos_IDS", StringComparison.CurrentCulture))
                {
                    int indexRow = flowsheet.Value.Rows.FindIndex(x =>
                        x.Parameter.Contains(measKey, StringComparison.CurrentCultureIgnoreCase));
                    if (indexRow >= 0)
                    {
                        flowsheet.Value.Rows.InsertRange(indexRow + 1, flowRows);
                    }
                }
            }
        }

        private static void InsertRtosIdsInstRows(string measKey, InstanceRows instanceRows)
        {
            if (!instanceRows.Any())
            {
                return;
            }

            foreach (KeyValuePair<string, InstanceSheet> instsheet in TestProgram.IgxlWorkBk.InsSheets)
            {
                if (instsheet.Key.Contains("TestInst_Rtos", StringComparison.CurrentCulture))
                {
                    int indexRow = instsheet.Value.Rows.FindIndex(x =>
                        x.TestName.Contains(measKey, StringComparison.CurrentCultureIgnoreCase));
                    if (indexRow >= 0)
                    {
                        instsheet.Value.Rows.InsertRange(indexRow + 1, instanceRows);
                    }
                }
            }
        }
    }
}
