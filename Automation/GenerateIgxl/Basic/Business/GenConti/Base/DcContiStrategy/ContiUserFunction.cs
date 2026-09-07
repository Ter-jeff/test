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
    public class ContiUserFunction : ContiBase
    {
        public ContiUserFunction(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
        {
        }

        public override List<FlowRow> GenerateFlowRows()
        {
            var flowRows = new List<FlowRow>();
            var flowRow = new FlowRow();
            foreach (KeyValuePair<string, string> condition in DcTestContiRow.ConditionDict)
            {
                if (condition.Key.Equals("Opcode", StringComparison.OrdinalIgnoreCase) &&
                    condition.Value.Equals("Test", StringComparison.OrdinalIgnoreCase))
                {
                    flowRow.Opcode = condition.Value;
                }
            }
            flowRow.Job = string.Join(",", DcTestContiRow.JobNameList);
            flowRow.Parameter = DcTestContiRow.InstanceName;
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

                var grpTNameLimit = DcTestContiRow.Limits
                    .Where(p => !string.IsNullOrWhiteSpace(p.LimitHeader))
                    .ToList();

                int added = AddUseLimitRows(grpTNameLimit, flowRow.Parameter, flowRows);

            }
            return flowRows;
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {
            return new List<InstanceRow>();
        }

        public static int AddUseLimitRows(IEnumerable<DcTestContiSheetLimit> grpTNameLimit, string parameter, IList<FlowRow> flowRows)
        {
            if (grpTNameLimit == null || flowRows == null)
            {
                return 0;
            }

            int added = 0;

            foreach (DcTestContiSheetLimit limitRow in grpTNameLimit)
            {
                if (limitRow == null)
                {
                    continue;
                }

                string paraUserLimit = parameter + "_" + limitRow.LimitStage;
                var row = new FlowRow
                {
                    Opcode = OpCode.UseLimit,
                    Parameter = paraUserLimit,
                    TName = limitRow.LimitHeader
                };

                string highUnit = "A";
                string highScale = "m";
                string lowUnit = limitRow.LimitUnit;
                string lowScale = "";

                row.LoLim = limitRow.LimitValue;
                row.HiLim = limitRow.HiLimitValue;
                row.Scale = string.IsNullOrEmpty(lowScale) ? highScale : lowScale;
                row.Units = string.IsNullOrEmpty(lowUnit) ? highUnit : lowUnit;

                flowRows.Add(row);
                added++;
            }

            return added;
        }
    }
}
