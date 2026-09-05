using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

namespace RfLib.Dvdc.GenTemplate.Bussiness
{
    internal sealed partial class ChannelMapPinResolver(List<ChannelMapSheet> channelMapSheets)
    {
        [GeneratedRegex("::")]
        private static partial Regex MyRegex19();

        public List<ChannelMapSheet> MultiChannelMap { get; } = channelMapSheets;

        public string SearchPinInChannelMap(MeasPin measPin)
        {
            var pins = new List<string> { measPin.PinName };
            if (MultiChannelMap != null && MultiChannelMap.Count != 0)
            {

                if (measPin.PinName.Contains("::"))
                {
                    pins = [.. MyRegex19().Split(measPin.PinName)];
                }

                for (int i = 0; i < pins.Count; i++)
                {
                    // 20190819 channel map -> multi channel map by Kimi
                    var related_Channelpins = MultiChannelMap.SelectMany(c => c.Rows).Where
                        (p => p.DeviceUnderTestPinName.Contains(pins[i], StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (related_Channelpins.Count == 0)
                    {
                        string outString = string.Format("{0} Not Found in Channelmap", pins[i]);
                        Response.Report(outString, percentage: Convert.ToInt32(35));
                    }
                    else
                    {
                        if (measPin.MeasType.EqualsIgnoreCase(MeasType.MeasF))
                        {
                            IEnumerable<ChannelMapRow> related_Channelpins_io = related_Channelpins.Where(rc => rc.Type == "I/O");
                            if (related_Channelpins_io.Any())
                            {
                                pins[i] = related_Channelpins_io.First().DeviceUnderTestPinName;
                                continue;
                            }
                        }

                        var related_Channelpins_uvi80 = MultiChannelMap.SelectMany(c => c.Rows).Where
                            (p => p.DeviceUnderTestPinName.Contains(pins[i] + "_uvi80", StringComparison.OrdinalIgnoreCase)).ToList();
                        var related_Channelpins_dc30 = MultiChannelMap.SelectMany(c => c.Rows).Where
                            (p => p.DeviceUnderTestPinName.Contains(pins[i] + "_dc30", StringComparison.OrdinalIgnoreCase)).ToList();
                        var related_Channelpins_dcvs = MultiChannelMap.SelectMany(c => c.Rows).Where
                            (p => p.DeviceUnderTestPinName.Contains(pins[i], StringComparison.OrdinalIgnoreCase) &&
                                  p.Type.Contains("DCVS")
                            ).ToList();

                        if (measPin.MeasType.EqualsIgnoreCase(MeasType.MeasVdm))
                        {
                            if (
                                related_Channelpins_uvi80.Exists(
                                    p => p.Type.EqualsIgnoreCase("DCDiffMeter")))
                            {
                                pins[i] =
                                    related_Channelpins_uvi80.FirstOrDefault(
                                        p => p.Type.EqualsIgnoreCase("DCDiffMeter"))!
                                        .DeviceUnderTestPinName;
                            }
                            else if (
                              related_Channelpins.Exists(
                                  p => p.Type.EqualsIgnoreCase("DCDiffMeter")))
                            {
                                pins[i] =
                                    related_Channelpins.FirstOrDefault(
                                        p => p.Type.EqualsIgnoreCase("DCDiffMeter"))!
                                        .DeviceUnderTestPinName;
                            }
                            else
                            {
                                string outString = string.Format("{0} Not Found DCDiffMeter Type in Channelmap", pins[i]);
                                //Response.Report(outString, percentage: Convert.ToInt32(35));
                            }
                        }
                        else if (related_Channelpins_dc30.Count > 0)
                        {
                            pins[i] =
                                related_Channelpins_dc30[0].DeviceUnderTestPinName;

                        }
                        else if (related_Channelpins_uvi80.Count > 0)
                        {
                            pins[i] =
                                related_Channelpins_uvi80[0].DeviceUnderTestPinName;

                        }
                        else if (related_Channelpins_dcvs.Count > 0)
                        {
                            pins[i] =
                                related_Channelpins_dcvs[0].DeviceUnderTestPinName;

                        }
                    }
                }
            }
            else
            {
                Response.Report("Channelmap is Not Found", percentage: Convert.ToInt32(35));
            }
            return string.Join("::", pins);
        }
    }
}
