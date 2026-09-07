using System.Collections.Generic;

using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.GenT0TXProgram
{
    public class GenT0TXProgramMain
    {
        private Dictionary<string, SubFlowSheet> _mainProgramSubFlows = new Dictionary<string, SubFlowSheet>();
        private Dictionary<string, InstanceSheet> _mainProgramInstanceSheets = new Dictionary<string, InstanceSheet>();

        public GenT0TXProgramMain(Dictionary<string, SubFlowSheet> mainprogramSubFlows, Dictionary<string, InstanceSheet> mainprogramInstanceSheets)
        {
            _mainProgramSubFlows = mainprogramSubFlows;
            _mainProgramInstanceSheets = mainprogramInstanceSheets;
        }

        public void Workflow()
        {
            List<SubFlowSheet> t0txSubflows = ModifySubflows();
            foreach (SubFlowSheet flow in t0txSubflows)
            {
                TestProgram.T0TxIgxlWorkBk.AddSubFlowSheet(FolderStructure.DirT0Tx, flow);
            }

            List<InstanceSheet> t0txInstanceSheets = ModifyInstances();
            foreach (InstanceSheet instanceSheet in t0txInstanceSheets)
            {
                TestProgram.T0TxIgxlWorkBk.AddInsSheet(FolderStructure.DirT0Tx, instanceSheet);
            }
        }

        public List<SubFlowSheet> ModifySubflows()
        {
            var results = new List<SubFlowSheet>();
            foreach (KeyValuePair<string, SubFlowSheet> flow in _mainProgramSubFlows)
            {
                var t0txFlow = new SubFlowSheet(flow.Value.Name);
                foreach (FlowRow flowRow in flow.Value.Rows)
                {
                    var t0txFlowRow = new FlowRow(flowRow);
                    if (!string.IsNullOrEmpty(t0txFlowRow.Enable))
                    {
                        t0txFlowRow.Enable = t0txFlowRow.Enable.Replace("CP1", "FT1");
                        t0txFlowRow.Enable = t0txFlowRow.Enable.Replace("CP2", "FT2");
                    }
                    if (!string.IsNullOrEmpty(t0txFlowRow.Job))
                    {
                        t0txFlowRow.Job = t0txFlowRow.Job.Replace("CP1", "FT1");
                        t0txFlowRow.Job = t0txFlowRow.Job.Replace("CP2", "FT2");
                    }
                    t0txFlow.AddRow(t0txFlowRow);
                }
                results.Add(t0txFlow);
            }
            return results;
        }

        public List<InstanceSheet> ModifyInstances()
        {
            var results = new List<InstanceSheet>();
            foreach (KeyValuePair<string, InstanceSheet> instanceSheet in _mainProgramInstanceSheets)
            {
                if (!instanceSheet.Value.Rows.Exists(x => !string.IsNullOrEmpty(x.GetArgument("isHarvesting"))))
                {
                    results.Add(instanceSheet.Value);
                    continue;
                }
                var t0txInstanceSheet = new InstanceSheet(instanceSheet.Value.Name + "_T0TX");
                foreach (InstanceRow instanceRow in instanceSheet.Value.Rows)
                {
                    var t0txInstanceRow = new InstanceRow(instanceRow);
                    if (string.IsNullOrEmpty(t0txInstanceRow.GetArgument("isHarvesting")))
                    {
                        t0txInstanceSheet.AddRow(t0txInstanceRow);
                    }
                    else
                    {
                        t0txInstanceRow.SetArgument("isHarvesting", "FALSE");
                        t0txInstanceSheet.AddRow(t0txInstanceRow);
                    }
                }
                results.Add(t0txInstanceSheet);
            }
            return results;
        }
    }
}
