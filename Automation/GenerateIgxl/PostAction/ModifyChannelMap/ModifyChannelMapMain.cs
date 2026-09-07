using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Singleton;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.ModifyChannelMap
{
    public class ModifyChannelMapMain
    {
        public void WorkFlow()
        {
            Dictionary<string, ChannelMapSheet> channelMapSheets = TestProgram.IgxlWorkBk.ChannelMapSheets;
            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            MultiTestSettingSheetsSingleton testsettings = MultiTestSettingSheetsSingleton.Instance();
            List<string> powerPinList = testsettings.PowerPinList;
            if (channelMapSheets == null || pinMap == null)
            {
                return;
            }

            #region Add extra DCVI, DCVS pins

            IEnumerable<IGrouping<string, string>> allPinswithType = TestProgram.IgxlWorkBk.ChannelMapSheets.SelectMany(sheet => sheet.Value.Rows).GroupBy(row => row.DeviceUnderTestPinName, row => row.Type);

            List<string> dcvidcvsPins = new List<string>();
            foreach (IGrouping<string, string> pinswithType in allPinswithType)
            {
                IEnumerable<string> dcvspins = pinswithType.Where(p => Regex.IsMatch(p, "DCVS", RegexOptions.IgnoreCase));
                IEnumerable<string> dcvipins = pinswithType.Where(p => Regex.IsMatch(p, "DCVI", RegexOptions.IgnoreCase));
                if (!powerPinList.Exists(x => x.Equals(pinswithType.Key, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (dcvspins.Any() && dcvipins.Any())
                {
                    dcvidcvsPins.Add(pinswithType.Key);
                }
            }


            foreach (string channelMapSheet in channelMapSheets.Keys)
            {
                IEnumerable<ChannelMapRow> chData = channelMapSheets[channelMapSheet].Rows.Where(row => dcvidcvsPins.Contains(row.DeviceUnderTestPinName));
                List<ChannelMapRow> newchRows = new List<ChannelMapRow>();
                List<ChannelMapRow> delchRows = new List<ChannelMapRow>();
                foreach (ChannelMapRow chpin in chData)
                {

                    if (chpin.Type.Contains("DCVS"))
                    {
                        newchRows.Add(new ChannelMapRow(chpin.DeviceUnderTestPinName + "_DCVI", "N/C"));
                        chpin.DeviceUnderTestPinName += "_DCVS";
                    }
                    else if (chpin.Type == "DCVI")
                    {
                        newchRows.Add(new ChannelMapRow(chpin.DeviceUnderTestPinName + "_DCVS", "N/C"));
                        chpin.DeviceUnderTestPinName += "_DCVI";
                    }
                    else if (chpin.Type == "N/C")
                    {
                        newchRows.Add(new ChannelMapRow(chpin.DeviceUnderTestPinName + "_DCVS", "N/C"));
                        newchRows.Add(new ChannelMapRow(chpin.DeviceUnderTestPinName + "_DCVI", "N/C"));
                        delchRows.Add(chpin);
                    }
                }
                channelMapSheets[channelMapSheet].AddRows(newchRows);
                foreach (ChannelMapRow row in delchRows)
                {
                    channelMapSheets[channelMapSheet].RemoveRow(row);
                }
            }

            #endregion

            #region Add extra N/C pins

            foreach (string channelMapSheet in channelMapSheets.Keys)
            {
                List<ChannelMapRow> chData = channelMapSheets[channelMapSheet].Rows;
                List<string> channelMapPins = chData.Select(channel => channel.DeviceUnderTestPinName.ToUpper()).ToList();
                List<string> channelMapNonExistPins = (from pin in pinMap.PinList where !channelMapPins.Contains(pin.PinName.ToUpper()) select pin.PinName).ToList();
                if (channelMapNonExistPins.Count > 0)
                {
                    foreach (string pin in channelMapNonExistPins)
                    {
                        var channelMapRow = new ChannelMapRow { DeviceUnderTestPinName = pin, Type = "N/C" };
                        channelMapSheets[channelMapSheet].AddRow(channelMapRow);
                    }
                }
            }

            #endregion
        }
    }
}
