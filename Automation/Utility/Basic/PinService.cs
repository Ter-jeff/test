using System;
using System.Collections.Generic;
using System.Linq;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.Utility.Basic
{
    public static class PinService
    {
        public static List<string> GetPinTypeFromChannelMaps(string pinName, List<string> powerPinList, Dictionary<string, ChannelMapSheet> channelMapSheets)
        {
            List<string> pinType = new List<string>();
            List<string> powerPins = powerPinList;
            if (!powerPins.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                return pinType;
            }

            if (channelMapSheets != null)
            {
                if (channelMapSheets.SelectMany(x => x.Value.Rows).ToList().Exists(y => y.DeviceUnderTestPinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                {
                    List<ChannelMapRow> allPins = channelMapSheets.SelectMany(x => x.Value.Rows).ToList().FindAll(y => y.DeviceUnderTestPinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
                    pinType.AddRange(allPins.Select(pin => pin.Type).ToList());
                }
            }
            return pinType.Distinct().ToList();
        }
    }
}
