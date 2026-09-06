using System.Collections.Generic;

namespace IgxlLib.IgxlBase
{
    public class PortSet : IgxlRow
    {
        public string PortName { get; set; } = string.Empty;
        public List<PortRow> PortRows { get; set; } = [];

        public PortSet()
        {
        }

        public PortSet(string portName)
        {
            PortName = portName;
        }

        public void AddPortRow(PortRow portRow)
        {
            PortRows.Add(portRow);
        }
    }
}
