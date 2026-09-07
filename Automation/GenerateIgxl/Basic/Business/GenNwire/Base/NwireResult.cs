using System.Collections.Generic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Base
{
    public class NwireResult
    {
        #region Properity

        public List<SubFlowSheet> NWireFlowSheets { get; set; } = new List<SubFlowSheet>();

        public List<InstanceRow> NWireInstanceRows { get; set; } = new List<InstanceRow>();

        public List<PortMapSheet> PortMapSheets { get; set; }

        public List<BinTableRow> BinTables { get; set; } = new List<BinTableRow>();

        #endregion
    }
}
