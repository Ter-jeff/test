using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

namespace Automation.GenerateIgxl.Scan.Harvest.Flow
{
    public class ScanNonBinCutInstance : NonBinCutInstanceBase
    {
        private const string AMtrtsnsRead = "A_MTRTSNS_Read";

        public override List<FlowRow> GetIfFlowRows(BinCutFinalInstanceRow dataRow, List<FlowRow> testFlowRows, List<FlowRow> limitRows, bool isApplyEnv = true)
        {
            FlowRow testFlowRow = testFlowRows.FirstOrDefault() ?? new FlowRow();
            var ifFlowRows = new List<FlowRow>();
            var flowRow = new FlowRow();
            string siteVar = dataRow == null ? "" : dataRow.BinCutInstanceRow.SiteVar;
            string callFlow = dataRow == null ? "" : dataRow.BinCutInstanceRow.CallFlow;
            string domain = dataRow == null ? "" : dataRow.Domain.ToUpper();
            string block = dataRow == null ? "" : dataRow.Block.ToUpper();
            ifFlowRows.AddRange(testFlowRows);
            if (testFlowRow.IsSsn)
            {
                ifFlowRows.Add(new FlowRow { Opcode = OpCode.BinTable, Parameter = $"Bin_{domain}_{block}_SSN_INIT_Fail" });
            }
            if (limitRows != null)
            {
                ifFlowRows.AddRange(limitRows);
            }

            if (!string.IsNullOrEmpty(siteVar))
            {
                if (!string.IsNullOrEmpty(testFlowRow.DeviceName) || siteVar.Contains("&&") || siteVar.Contains("||"))
                {
                    ifFlowRows.Clear();
                    ifFlowRows.Add(FlowRow.GenIfCondition(siteVar, testFlowRow.Job));
                    ifFlowRows.AddRange(testFlowRows);
                    if (testFlowRow.IsSsn)
                    {
                        ifFlowRows.Add(new FlowRow { Opcode = OpCode.BinTable, Parameter = $"Bin_{domain}_{block}_SSN_INIT_Fail" });
                    }
                    if (limitRows != null)
                    {
                        ifFlowRows.AddRange(limitRows);
                    }

                    if (!string.IsNullOrEmpty(callFlow))
                    {
                        ifFlowRows.Add(GenCallFlowRow(callFlow, flowRow));
                    }

                    ifFlowRows.Add(FlowRow.GenEndIf(testFlowRow.Job));
                    return ifFlowRows;
                }

                if (string.IsNullOrEmpty(testFlowRow.DeviceName))
                {
                    AddCondition(siteVar, ref testFlowRow);
                }
            }

            if (!string.IsNullOrEmpty(callFlow))
            {
                if (!string.IsNullOrEmpty(testFlowRow.DeviceName))
                {
                    var ifFlowRow = new FlowRow
                    {
                        Enable = AMtrtsnsRead,
                        Job = flowRow.Job,
                        Opcode = "If",
                        Parameter = siteVar
                    };
                    ifFlowRows.Add(ifFlowRow);

                    ifFlowRows.Add(GenCallFlowRow(callFlow, flowRow));

                    var endIfFlowRow = new FlowRow { Enable = AMtrtsnsRead, Job = flowRow.Job, Opcode = "EndIf" };
                    ifFlowRows.Add(endIfFlowRow);
                }
                else
                {
                    ifFlowRows.Add(GenCallFlowRow(callFlow, flowRow));
                }
            }

            return ifFlowRows;
        }

        private FlowRow GenCallFlowRow(string callFlow, FlowRow flowRow)
        {
            var callFlowRow = new FlowRow
            {
                Enable = AMtrtsnsRead,
                Job = flowRow.Job,
                Opcode = "call",
                Parameter = callFlow
            };
            return callFlowRow;
        }
    }
}
