using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Reader;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public class ContiOpcode : ContiBase
    {
        public ContiOpcode(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
        {
        }

        public override List<FlowRow> GenerateFlowRows()
        {
            var flowRows = new List<FlowRow>();
            var flowRow = new FlowRow();
            List<string> instanceStr = DcTestContiRow.InstanceName.Split(':').ToList();
            flowRow.Job = string.Join(",", DcTestContiRow.JobNameList);
            flowRow.Opcode = instanceStr[0];
            if (instanceStr.Count > 1)
            {
                flowRow.Parameter = instanceStr[1];
            }
            flowRow.Enable = DcTestContiRow.EnableWord;

            if (!string.IsNullOrEmpty(flowRow.Opcode))
            {
                flowRows.Add(flowRow);

                if (!string.IsNullOrEmpty(DcTestContiRow.FailFlag.Trim()))
                {
                    var binTableRow = new FlowRow();
                    IEnumerable<string> flagList = DcTestContiRow.FailFlag.Replace(" ", "")
                                                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                .Select(x => Regex.Replace(x.Trim(), "^F_", "", RegexOptions.IgnoreCase));
                    binTableRow.Opcode = OpCode.BinTable;
                    binTableRow.Parameter = "Bin_" + string.Join("_", flagList);
                    flowRows.Add(binTableRow);
                }
            }

            return flowRows;
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {
            return new List<InstanceRow>();
        }
    }
}
