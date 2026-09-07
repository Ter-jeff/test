using System.Collections.Generic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.BistBira.Base
{
    public class BistIgxlResult
    {
        private readonly Dictionary<string, List<InstanceRow>> _dummyInstanceRows = new Dictionary<string, List<InstanceRow>>();

        public List<SubFlowSheet> FlowSheets { get; set; } = new List<SubFlowSheet>();
        public List<InstanceSheet> InstanceSheets { get; set; } = new List<InstanceSheet>();
        public List<BinTableRow> BinTableRows { get; set; } = new List<BinTableRow>();

        public void AddDummyInstance(string sheetName, InstanceRow row)
        {
            if (_dummyInstanceRows.TryGetValue(sheetName, out List<InstanceRow> instanceRow))
            {
                instanceRow.Add(row);
            }
            else
            {
                List<InstanceRow> rowList = new List<InstanceRow> { row };
                _dummyInstanceRows.Add(sheetName, rowList);
            }
        }
    }
}
