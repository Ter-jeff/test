using System.Collections.Generic;

namespace IgxlLib.IgxlBase
{
    public class ChannelMapRow : IgxlRow
    {
        public string DeviceUnderTestPinName { get; set; } = string.Empty;
        public string DeviceUnderTestPackagePin { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public List<string> Sites { get; set; } = [];
        public string Comment { get; set; } = string.Empty;
        public string InstrumentType { get; set; } = string.Empty;

        public ChannelMapRow()
        {
        }

        public ChannelMapRow(string pinName, string type)
        {
            DeviceUnderTestPinName = pinName;
            Type = type;
        }
    }
}
