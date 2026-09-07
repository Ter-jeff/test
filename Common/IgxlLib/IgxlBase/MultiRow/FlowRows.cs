using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace IgxlLib.IgxlBase.MultiRow
{
    public class FlowRows : List<FlowRow>
    {
        public FlowRows()
        {
        }

        public FlowRows(IEnumerable<FlowRow> flowRows) : base(flowRows)
        {
        }

        public void SetTestNumber(Dictionary<string, SubFlowSheet> subFlowSheets, long startNumber, long step, long maxNum, ref int count, ref List<string> doneSheets)
        {
            foreach (FlowRow flowRow in this)
            {
                if (flowRow.Opcode.EqualsIgnoreCase(OpCode.Test) || flowRow.Opcode.EqualsIgnoreCase("test-defer-limits"))
                {
                    if (string.IsNullOrEmpty(flowRow.TNum))
                    {
                        long number = startNumber + (count * step);
                        flowRow.TNum = number.ToString(CultureInfo.InvariantCulture);
                        if (startNumber + ((count + 1) * step) < maxNum)
                        {
                            count++;
                        }
                    }
                }

                if (subFlowSheets.Values.Any(x => x.Name.EqualsIgnoreCase(flowRow.Parameter)))
                {
                    if (!doneSheets.Exists(x => x.EqualsIgnoreCase(flowRow.Parameter)))
                    {
                        string start = (startNumber + (count * step)).ToString(CultureInfo.InvariantCulture);
                        SubFlowSheet sheet = subFlowSheets.Values.First(x => x.Name.EqualsIgnoreCase(flowRow.Parameter));
                        var rows = new FlowRows();
                        rows.AddRange(sheet.Rows);
                        rows.SetTestNumber(subFlowSheets, startNumber, step, maxNum, ref count, ref doneSheets);
                        string end = (startNumber + (count * step)).ToString(CultureInfo.InvariantCulture);
                        flowRow.AddColumnA($"TNum {start} - {end}");
                        doneSheets.Add(sheet.Name);
                    }
                }
            }
        }
    }
}
