using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.Static;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.GenGroupFlowSheet
{
    internal class GroupFlowSheetMain
    {
        public List<MainFlowBase> WorkFlow(Dictionary<string, SubFlowSheet> subFlowSheets, List<MainFlowBase> mainFlows)
        {
            foreach (MainFlowBase mainFlow in mainFlows)
            {
                foreach (FlowSequence sequence in mainFlow.Sequences)
                {
                    if (subFlowSheets.Values.Any(x => x.Name.Equals(sequence.SubFlowName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        EnumGroupInMainFlow groupInMainFlow = subFlowSheets.Values.First(x => x.Name.Equals(sequence.SubFlowName, StringComparison.CurrentCultureIgnoreCase)).GroupNameInMainFlow;
                        if (groupInMainFlow == EnumGroupInMainFlow.FlowScanSa)
                        {
                            sequence.GroupNameInMainFlow = EnumGroupInMainFlow.FlowScanSa;
                        }
                    }
                }

                var groups = mainFlow.Sequences.Where(x => !string.IsNullOrEmpty(x.SubFlowName)).ChunkBy(x => x.GroupNameInMainFlow.ToString()).ToList(); //remove empty cell
                int cnt = 0;
                for (int i = 0; i < groups.Count; i++)
                {
                    IGrouping<string, FlowSequence> group = groups.ElementAt(i);
                    if (group.Key == nameof(EnumGroupInMainFlow.None) || string.IsNullOrEmpty(group.Key))
                    {
                        continue;
                    }

                    string groupSheetName = cnt == 0
                        ? Combination.CombineByUnderLine(group.Key, mainFlow.JobName)
                        : Combination.CombineByUnderLine(group.Key, mainFlow.JobName) + "_" + cnt;
                    //Set for main flow
                    foreach (FlowSequence item in group)
                    {
                        item.GroupSheetName = groupSheetName;
                    }

                    cnt++;
                    GenFlow(groupSheetName, group.ToList(), mainFlow.JobName);
                }
            }
            return mainFlows;
        }

        private void GenFlow(string sheetName, List<FlowSequence> flowSequences, string jobName)
        {
            var flow = new SubFlowSheet(sheetName) { JobNames = new List<string> { jobName } };
            foreach (FlowSequence flowSequence in flowSequences)
            {
                var flowRow = new FlowRow
                {
                    Opcode = TestProgram.IgxlWorkBk.SubFlowSheets.Any(p => p.Value.Name.Equals(flowSequence.SubFlowName, StringComparison.OrdinalIgnoreCase))
                    ? OpCode.Call : OpCode.Nop,
                    Parameter = flowSequence.SubFlowName,
                    Enable = flowSequence.Enable
                };
                flow.AddRow(flowRow);
            }
            flow.AddReturnRow();
            if (flowSequences.First().GroupNameInMainFlow == EnumGroupInMainFlow.FlowScanSa)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirNonBinCut, flow);
            }
            else
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirMain, flow);
            }
        }
    }
}
