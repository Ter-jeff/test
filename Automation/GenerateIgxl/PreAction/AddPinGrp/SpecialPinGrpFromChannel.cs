using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Basic;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using TestPlanLib.DataStruct;

namespace Automation.GenerateIgxl.PreAction.AddPinGrp
{
    internal class SpecialPinGrpFromChannel
    {
        private readonly MultiTestSettingSheetsSingleton _testSettings = MultiTestSettingSheetsSingleton.Instance();

        public void WorkFlow()
        {
            List<PinGroup> newGroups = CreateDcviDcvsPinGrp();
            foreach (PinGroup group in newGroups)
            {
                if (group.PinList.Count == 0)
                {
                    continue;
                }

                AddPinGrp(group);
            }

            PinMapSheet pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap == null)
            {
                return;
            }

            PinGroup pinGroup = new PinGroup("All_Digital_Disconnect_Continuity", PinMapConst.TypeIo);
            List<Pin> pinList = pinMap.GetAllDigitalDisconnectContinuityPins();
            pinGroup.AddPins(pinList);
            if (pinGroup.PinList.Count != 0)
            {
                pinMap.AddGroup(pinGroup);
            }

            List<PinGroup> dcviGroups = pinMap.GenDcviGroup();
            foreach (PinGroup group in dcviGroups)
            {
                pinMap.AddGroup(group);
            }

            pinMap.GenDcvsGroup();

        }

        private List<PinGroup> CreateDcviDcvsPinGrp()
        {
            var allPinGroups = new List<PinGroup>();
            var dcvsPowerGroup = new List<PinGroup>();
            var dcviPowerGroup = new List<PinGroup>();
            List<Pin> allPowerPin = TestProgram.IgxlWorkBk.PinMapPair.Value.GetPowerPins();
            var powerPinGroups = new Dictionary<string, List<string>>();
            PowerInfoSheet powerInfoSheet = TestPlanStatic.PowerInfoSheet;
            var powerInfos = new List<PowerInfoRow>(powerInfoSheet.Rows);
            var diffPinTypePinGroup = new List<string>();
            SplitDiffPinTypePowerInfos(powerInfos, diffPinTypePinGroup);
            if (powerInfoSheet != null)
            {
                powerPinGroups = powerInfos.GroupBy(p => p.Chiplet).ToDictionary(p => p.Key, p => p.Select(q => q.PinName).ToList());

                CreatePowerPinGroups(powerPinGroups, diffPinTypePinGroup, dcvsPowerGroup, dcviPowerGroup);
            }


            foreach (string sheet in TestProgram.IgxlWorkBk.ChannelMapSheets.Keys)
            {
                ChannelMapSheet channelMapSheet = TestProgram.IgxlWorkBk.ChannelMapSheets[sheet];
                foreach (ChannelMapRow chRow in channelMapSheet.Rows)
                {
                    if (!powerPinGroups.Values.SelectMany(p => p).ToList().Exists(p => p.Equals(chRow.DeviceUnderTestPinName)))
                    {
                        continue;
                    }

                    if (Regex.IsMatch(chRow.Type, "DCVI", RegexOptions.IgnoreCase))
                    {
                        if (allPowerPin.Exists(p => p.PinName.Equals(chRow.DeviceUnderTestPinName)) || diffPinTypePinGroup.Exists(x => x.Equals(chRow.DeviceUnderTestPinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            AddPinToMatchingPowerGroup(powerPinGroups, dcviPowerGroup, chRow.DeviceUnderTestPinName);
                        }
                    }
                    if (Regex.IsMatch(chRow.Type, "DCVS", RegexOptions.IgnoreCase))
                    {
                        if (allPowerPin.Exists(p => p.PinName.Equals(chRow.DeviceUnderTestPinName)) || diffPinTypePinGroup.Exists(x => x.Equals(chRow.DeviceUnderTestPinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            AddPinToMatchingPowerGroup(powerPinGroups, dcvsPowerGroup, chRow.DeviceUnderTestPinName);
                        }
                    }
                }
            }
            allPinGroups.AddRange(dcviPowerGroup);
            allPinGroups.AddRange(dcvsPowerGroup);
            if (dcvsPowerGroup.Any())
            {
                LocalSpecs.ExistPowerPinListDcvs = true;
            }
            if (dcviPowerGroup.Any())
            {
                LocalSpecs.ExistPowerPinListDcvi = true;
            }
            return allPinGroups;
        }

        private void SplitDiffPinTypePowerInfos(List<PowerInfoRow> powerInfos, List<string> diffPinTypePinGroup)
        {
            for (int i = 0; i < powerInfos.Count; i++)
            {
                string pinName = powerInfos[i].PinName;
                List<string> allPinType = PinService.GetPinTypeFromChannelMaps(pinName, _testSettings.PowerPinList, TestProgram.IgxlWorkBk.ChannelMapSheets);

                if (allPinType.Count > 1)
                {
                    IEnumerable<string> dcvsPins = allPinType.Where(p => Regex.IsMatch(p, "DCVS", RegexOptions.IgnoreCase));
                    IEnumerable<string> dcviPins = allPinType.Where(p => Regex.IsMatch(p, "DCVI", RegexOptions.IgnoreCase));
                    if (dcvsPins.Any() && dcviPins.Any())
                    {
                        if (allPinType.Count >= 2)
                        {
                            PowerInfoRow newDcvsPin = powerInfos[i].Copy();
                            newDcvsPin.PinName = pinName + "_DCVS";
                            powerInfos.Insert(i, newDcvsPin);
                            PowerInfoRow newDcviPin = powerInfos[i].Copy();
                            newDcviPin.PinName = pinName + "_DCVI";
                            powerInfos.Insert(i, newDcviPin);

                            powerInfos.RemoveAt(i + 2);
                            diffPinTypePinGroup.Add(pinName);
                        }
                    }
                }
            }
        }

        private void CreatePowerPinGroups(Dictionary<string, List<string>> powerPinGroups, List<string> diffPinTypePinGroup, List<PinGroup> dcvsPowerGroup, List<PinGroup> dcviPowerGroup)
        {
            foreach (KeyValuePair<string, List<string>> powerPinGroup in powerPinGroups)
            {
                var pins = new List<string>();
                if (diffPinTypePinGroup.Exists(x => x.Equals(powerPinGroup.Key)))
                {
                    pins.AddRange(powerPinGroup.Value);
                }
                else
                {
                    pins.Add(powerPinGroup.Key);
                }
                foreach (string pin in pins)
                {
                    if (string.IsNullOrEmpty(powerPinGroup.Key))
                    {
                        dcvsPowerGroup.Add(new PinGroup(CommonConst.DcvsPower, PinMapConst.TypePower));
                        dcviPowerGroup.Add(new PinGroup(CommonConst.DcviPower, PinMapConst.TypePower));

                    }
                    else
                    {
                        dcvsPowerGroup.Add(new PinGroup(CommonConst.DcvsPower + "_" + pin, PinMapConst.TypePower));
                        dcviPowerGroup.Add(new PinGroup(CommonConst.DcviPower + "_" + pin, PinMapConst.TypePower));

                    }
                }
            }
        }

        private void AddPinToMatchingPowerGroup(Dictionary<string, List<string>> powerPinGroups, List<PinGroup> targetPowerGroup, string dutPinName)
        {
            foreach (KeyValuePair<string, List<string>> powerPinGroup in powerPinGroups)
            {
                if (powerPinGroup.Value.Contains(dutPinName))
                {
                    string postFixed = string.IsNullOrEmpty(powerPinGroup.Key)
                        ? ""
                        : "_" + powerPinGroup.Key;
                    PinGroup group = targetPowerGroup.FirstOrDefault(p => Regex.IsMatch(p.PinName, postFixed, RegexOptions.IgnoreCase));
                    group?.AddPin(dutPinName);
                }
            }
        }

        private void AddPinGrp(PinGroup pinGroup)
        {
            if (pinGroup.PinList.Count > 0)
            {
                TestProgram.IgxlWorkBk.PinMapPair.Value.AddRow(pinGroup);
            }
        }
    }
}
