using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow.Business
{
    internal class DeferredBinOutProcessor
    {
        internal Dictionary<string, HashSet<string>> GetDeferredBinoutJobs(MainFlowSheet mainFlowSheet)
        {
            Dictionary<string, HashSet<string>> deferredBinoutJobs = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (MainFlowBase mainFlow in mainFlowSheet.Rows)
            {
                IEnumerable<FlowSequenceNew> deferredBinoutFlows = mainFlow.SequencesNew.Where(x => x.Module.Equals("Defer", StringComparison.OrdinalIgnoreCase) && x.Enable);
                foreach (FlowSequenceNew deferredBinoutFlow in deferredBinoutFlows)
                {
                    if (!deferredBinoutJobs.ContainsKey(deferredBinoutFlow.SheetName))
                    {
                        deferredBinoutJobs.Add(deferredBinoutFlow.SheetName, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
                    }
                    deferredBinoutFlow.GetJobList().ForEach(x => deferredBinoutJobs[deferredBinoutFlow.SheetName].Add(x));
                }
            }
            return deferredBinoutJobs;
        }

        internal Dictionary<string, List<FlowSequenceNew>> GetDeferredBinoutMaps(MainFlowSheet mainFlowSheet, Dictionary<string, HashSet<string>> deferredBinoutJobs)
        {
            Dictionary<string, List<FlowSequenceNew>> deferredBinoutMaps = new Dictionary<string, List<FlowSequenceNew>>(StringComparer.OrdinalIgnoreCase);
            List<FlowSequenceNew> deferredSourceFlows = mainFlowSheet.Rows.FirstOrDefault().SequencesNew.Where(x => x.OptionDict.ContainsKey("Defer")).ToList();
            foreach (string deferredBinoutFlow in deferredBinoutJobs.Keys)
            {
                deferredBinoutMaps[deferredBinoutFlow] = deferredSourceFlows.Where(x => x.OptionDict["Defer"].EqualsIgnoreCase(deferredBinoutFlow)).ToList();
            }
            return deferredBinoutMaps;
        }

        internal Dictionary<string, SubFlowSheet> GetDeferredBinoutFlows(
            InstanceSheet commonInstanceSheet,
            Dictionary<string, List<SubFlowSheet>> deferredBinoutMapFlows,
            Dictionary<string, HashSet<string>> deferredBinoutJobs,
            IEnumerable<string> allJobs)
        {
            var deferredSubFlowSheets = new Dictionary<string, SubFlowSheet>();
            Dictionary<string, List<BinFlowItem>> deferredFlowBinTables = GetDeferredBinFlows(deferredBinoutMapFlows, deferredBinoutJobs, allJobs);

            foreach (KeyValuePair<string, List<BinFlowItem>> deferredFlow in deferredFlowBinTables)
            {
                SubFlowSheet sheet = new SubFlowSheet(deferredFlow.Key, deferredFlow.Key);
                sheet.AddStartRows(sheet.Name);
                if (!deferredFlow.Value.Any())
                {
                    continue;
                }
                foreach (BinFlowItem binFlowItem in deferredFlow.Value)
                {
                    if (binFlowItem.IfRow != null)
                    {
                        sheet.AddRow(binFlowItem.IfRow);
                    }
                    sheet.AddRow(binFlowItem.BinTableRow);
                    if (binFlowItem.EndIfRow != null)
                    {
                        sheet.AddRow(binFlowItem.EndIfRow);
                    }
                }
                commonInstanceSheet?.AddHeaderFooter(sheet.Name);
                sheet.AddEndRows(sheet.Name);
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirCommon, sheet);
                string name = Path.Combine(FolderStructure.DirCommon, sheet.Name);
                deferredSubFlowSheets[name] = sheet;
            }
            return deferredSubFlowSheets;
        }

        internal Dictionary<string, List<BinFlowItem>> GetDeferredBinFlows(
            Dictionary<string, List<SubFlowSheet>> deferredBinoutMapFlows,
            Dictionary<string, HashSet<string>> deferredBinoutJobs,
            IEnumerable<string> allJobs)
        {
            Dictionary<string, List<BinFlowItem>> deferredBinFlows = new Dictionary<string, List<BinFlowItem>>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, HashSet<string>> deferredBinoutJob in deferredBinoutJobs)
            {
                string deferredBinoutFlowName = deferredBinoutJob.Key;
                HashSet<string> deferredBinoutFlowJobs = deferredBinoutJob.Value;
                if (!deferredBinoutMapFlows.TryGetValue(deferredBinoutFlowName, out List<SubFlowSheet> deferredSourceFlows))
                {
                    continue;
                }

                List<FlowRow> flowRows = deferredSourceFlows.SelectMany(x => x.Rows).Where(x => !x.IsBackup).ToList();
                IEnumerable<BinFlowItem> binFlowItems = CreateBinFlowItems(flowRows);
                List<string> nonAvalibleJobs = allJobs.Except(deferredBinoutFlowJobs).ToList();

                foreach (FlowRow binFlow in binFlowItems.Select(x => x.BinTableRow))
                {
                    if (binFlow.Opcode == OpCode.BinTable)
                    {
                        List<string> jobs = new List<string>();
                        jobs.AddRange(nonAvalibleJobs);
                        if (jobs.Any())
                        {
                            if (binFlow.Parameter.StartsWith("Bin_HIP", StringComparison.OrdinalIgnoreCase))
                            {
                                string hipEnableWord = string.Join("||", jobs.Select(x => $"Prod_{x}"));
                                if (string.IsNullOrEmpty(binFlow.Enable))
                                {
                                    binFlow.Enable = hipEnableWord;
                                }
                                else
                                {
                                    binFlow.Enable = $"({binFlow.Enable})&&({hipEnableWord})";
                                }
                            }
                            else
                            {
                                binFlow.Job = string.Join(",", jobs);
                            }
                        }
                        else
                        {
                            binFlow.Opcode = OpCode.Nop;
                        }
                    }
                }
                deferredBinFlows[deferredBinoutFlowName] = new List<BinFlowItem>();
                foreach (BinFlowItem item in binFlowItems)
                {
                    if (!ContainsSameItem(deferredBinFlows[deferredBinoutFlowName], item))
                    {
                        var newItem = new BinFlowItem(item.BinTableRow.Copy(), item.IfRow?.Copy() ?? null, item.EndIfRow?.Copy() ?? null);
                        newItem.BinTableRow.Enable = "";
                        newItem.BinTableRow.Job = string.Join(",", deferredBinoutFlowJobs);
                        deferredBinFlows[deferredBinoutFlowName].Add(newItem);
                    }
                }
            }
            return deferredBinFlows;
        }

        private static List<BinFlowItem> CreateBinFlowItems(
        IReadOnlyList<FlowRow> flowRows)
        {
            List<BinFlowItem> result = [];

            for (int index = 0; index < flowRows.Count; index++)
            {
                FlowRow current = flowRows[index];

                if (current.Opcode != OpCode.BinTable)
                {
                    continue;
                }

                FlowRow previous = index > 0
                    ? flowRows[index - 1]
                    : null;

                FlowRow next = index < flowRows.Count - 1
                    ? flowRows[index + 1]
                    : null;

                bool isWrapped =
                    previous != null &&
                    previous.Opcode == OpCode.If &&
                    next != null &&
                    next.Opcode == OpCode.EndIf;

                result.Add(new BinFlowItem(
                    current,
                    isWrapped ? previous : null,
                    isWrapped ? next : null));
            }

            return result;
        }

        private static bool ContainsSameItem(
            IEnumerable<BinFlowItem> existingItems,
            BinFlowItem candidate)
        {
            return existingItems.Any(existing =>
                AreSame(existing, candidate));
        }

        private static bool AreSame(
            BinFlowItem x,
            BinFlowItem y)
        {
            bool xIsWrapped =
                x.IfRow != null &&
                x.EndIfRow != null;

            bool yIsWrapped =
                y.IfRow != null &&
                y.EndIfRow != null;

            if (xIsWrapped != yIsWrapped)
            {
                return false;
            }

            if (!xIsWrapped)
            {
                return string.Equals(
                    x.BinTableRow.Parameter,
                    y.BinTableRow.Parameter,
                    StringComparison.OrdinalIgnoreCase);
            }

            return string.Equals(
                       x.IfRow.Parameter,
                       y.IfRow.Parameter,
                       StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(
                       x.BinTableRow.Parameter,
                       y.BinTableRow.Parameter,
                       StringComparison.OrdinalIgnoreCase);
        }
    }
    internal sealed record BinFlowItem(
        FlowRow BinTableRow,
        FlowRow IfRow,
        FlowRow EndIfRow);
}
