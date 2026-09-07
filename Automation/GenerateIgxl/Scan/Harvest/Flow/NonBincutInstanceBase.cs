using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.Scan.Harvest.Flow
{
    public abstract class NonBinCutInstanceBase
    {

        protected void AddCondition(string siteVar, ref FlowRow flowRow)
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

        public virtual List<FlowRow> GetIfFlowRows(BinCutFinalInstanceRow dataRow, List<FlowRow> testFlowRows, List<FlowRow> flowRows, bool isApplyEnv = true)
        {
            FlowRow testFlowRow = testFlowRows.FirstOrDefault() ?? new FlowRow();
            var ifFlowRows = new List<FlowRow>();
            var flowRow = new FlowRow();
            string siteVar = dataRow == null ? "" : dataRow.BinCutInstanceRow.SiteVar;
            ifFlowRows.AddRange(testFlowRows);
            if (flowRows != null)
            {
                if (isApplyEnv)
                {
                    flowRows.ForEach(x => x.Env = testFlowRow.Env);
                }

                ifFlowRows.AddRange(flowRows);
            }
            if (!string.IsNullOrEmpty(siteVar))
            {
                if (!string.IsNullOrEmpty(testFlowRow.DeviceName) || siteVar.Contains("&&") || siteVar.Contains("||"))
                {
                    ifFlowRows.Clear();
                    ifFlowRows.Add(FlowRow.GenIfCondition(siteVar, testFlowRow.Job));
                    ifFlowRows.AddRange(testFlowRows);
                    if (flowRows != null)
                    {
                        ifFlowRows.AddRange(flowRows);
                    }

                    ifFlowRows.Add(FlowRow.GenEndIf(testFlowRow.Job));
                    return ifFlowRows;
                }

                if (string.IsNullOrEmpty(testFlowRow.DeviceName))
                {
                    AddCondition(siteVar, ref testFlowRow);
                }
            }
            return ifFlowRows;
        }
    }
}
