using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.PreAction.AddPinGrp;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility.Tester;

using LogLib.Static;

namespace Automation.GenerateIgxl.PreAction.GenChannelMap
{
    public class ChannelMapMain
    {
        private readonly Dictionary<string, TesterConfig> _testerConfigDic;

        public ChannelMapMain(string deviceType)
        {
            string file = Path.Combine(AppContext.BaseDirectory, "Config", "Tester", "TesterConfig_" + deviceType + ".xml");
            if (!File.Exists(file))
            {
                file = Path.Combine(AppContext.BaseDirectory, "Config", "Tester", "TesterConfig_Default.xml");
            }

            _testerConfigDic = TesterConfigReader.GetTesterConfigs(file);
        }

        public List<ChannelMapSheet> WorkFlow(string folder)
        {
            var channelMapSheets = new List<ChannelMapSheet>();
            string[] files = Directory.GetFiles(folder, "*.txt");
            if (!files.Any())
            {
                return channelMapSheets;
            }

            var channelSheets = files.Where(x => Regex.IsMatch(Path.GetFileName(x), "^Channel*", RegexOptions.IgnoreCase)).ToList();
            if (!channelSheets.Any())
            {
                return channelMapSheets;
            }

            foreach (string sheet in channelSheets)
            {
                ChannelMapSheet channelMapSheet = ReadChanMapSheet.GetSheet(sheet, _testerConfigDic);
                channelMapSheets.Add(channelMapSheet);
            }

            return channelMapSheets;
        }

        public void WorkFlow(List<ChannelMapSheet> channelSheets)
        {
            if (channelSheets.Count == 0)
            {
                Response.Report("Missing ChannelMap sheet in test plan file", EnumMessageLevel.Error, 10);
            }

            foreach (ChannelMapSheet channelMapSheet in channelSheets)
            {
                TestProgram.IgxlWorkBk.AddChannelMapSheet(FolderStructure.DirChannelMap, channelMapSheet);
            }
        }

        public void ModifyPinMapByChannelMap()
        {
            ModifyPinMap();

            var specialPinGrpFromChannel = new SpecialPinGrpFromChannel();
            specialPinGrpFromChannel.WorkFlow();
        }

        private void ModifyPinMap()
        {
            if (TestProgram.IgxlWorkBk.PinMapPair.Value != null)
            {
                PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
                var pinList = pinMap.PinList.Select(x => x.PinName).ToList();
                List<string> powerPins = MultiTestSettingSheetsSingleton.Instance().PowerPinList;
                foreach (string pinName in pinList)
                {
                    Pin pin = pinMap.PinList.FirstOrDefault(x => x.PinName.Equals(pinName));
                    int index = pinMap.PinList.Select((p, i) => new { Pin = p, Index = i }).FirstOrDefault(x => x.Pin.PinName.Equals(pinName))?.Index ?? -1;
                    if (TestProgram.IgxlWorkBk.ChannelMapSheets != null)
                    {
                        if (TestProgram.IgxlWorkBk.ChannelMapSheets.SelectMany(x => x.Value.Rows).ToList()
                            .Exists(y => y.DeviceUnderTestPinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            List<ChannelMapRow> channelMapRows = TestProgram.IgxlWorkBk.ChannelMapSheets.SelectMany(x => x.Value.Rows).ToList()
                                .FindAll(y => y.DeviceUnderTestPinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
                            var channelTypes = channelMapRows.Select(x => x.Type).Distinct().ToList();
                            string instrumentType = channelMapRows.FirstOrDefault()?.InstrumentType;
                            string channelType = channelMapRows.FirstOrDefault()?.Type;
                            if (channelTypes.Count > 1)
                            {
                                IEnumerable<string> dcvsPins = channelTypes.Where(p =>
                                    Regex.IsMatch(p, "DCVS", RegexOptions.IgnoreCase));
                                IEnumerable<string> dcviPins = channelTypes.Where(p =>
                                    Regex.IsMatch(p, "DCVI", RegexOptions.IgnoreCase));
                                if (dcvsPins.Any() && dcviPins.Any() && powerPins.Exists(x =>
                                        x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    if (channelTypes.Count >= 2)
                                    {
                                        if (pin != null)
                                        {
                                            pin.InstrumentType = channelMapRows.FirstOrDefault()?.InstrumentType;
                                            var newPinGroup = new PinGroup(pinName, PinMapConst.TypePower);
                                            Pin newDcvsPin = pin.Copy();
                                            newDcvsPin.PinName = pin.PinName + "_DCVS";
                                            newDcvsPin.ChannelType = "DCVS";
                                            newPinGroup.AddPin(newDcvsPin);
                                            pinMap.InsertPinAt(index, newDcvsPin);
                                            Pin newDcviPin = pin.Copy();
                                            newDcviPin.PinName = pin.PinName + "_DCVI";
                                            newDcviPin.ChannelType = "DCVI";
                                            pinMap.InsertPinAt(index, newDcviPin);
                                            pinMap.RemovePinAt(index + 2);
                                            newPinGroup.AddPin(newDcviPin);
                                            TestProgram.IgxlWorkBk.PinMapPair.Value.InsertGroup(0, newPinGroup);
                                        }
                                    }
                                }
                            }

                            if (pin != null)
                            {
                                pin.ChannelType = channelType;
                                pin.InstrumentType = instrumentType;
                            }
                        }
                    }
                }
            }
        }
    }
}
