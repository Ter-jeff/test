using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using RfLib.InstrumentSetup.InstrumentTypeData;

namespace RfLib.InstrumentSetup
{
    internal sealed partial class InstrumentTypeSelector(List<ChannelMapSheet> channelMapSheets)
    {
        private const string UltraPac80Source = "UltraPac80 Source";
        private const string UltraPac80Capture = "UltraPac80 Capture";
        private const string Gigadig = "Gigadig";
        private const string UVI80Vdm = "UVI80 Vdm";
        private const string UVI80Meter = "UVI80 Meter";
        private const string UVS256Meter = "UVS256 Meter";
        private const string UP1600 = "UP1600";
        private const string ErrorFrequency = "Error Frequency";
        private const string ErrorPinType = "Error Pin Type";

        [GeneratedRegex("hz", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        private readonly List<ChannelMapSheet> _multiChannelMap = channelMapSheets;

        public List<string> GetInstrumentType(MeasPin measPin, List<InstrumentTypeSetting> instrumentTypeSettings)
        {
            var type = new List<string>();
            if (instrumentTypeSettings.Count == 0)
            {
                return type;
            }
            // new setup for Loopback #33,#32
            if (measPin.MeasType.EqualsIgnoreCase("MeasLoopBack"))
            {
                type.AddRange(instrumentTypeSettings.FindAll(x => x.PinType.StartsWith("LoopBack")).Select(x => x.InstrumentType));
                return type;
            }
            string pinName = measPin.PinName.Contains("::") ? measPin.PinName.Split(["::"], StringSplitOptions.RemoveEmptyEntries)[0] : measPin.PinName;

            if (pinName.EndsWith("_LX"))
            {
                type.AddRange(instrumentTypeSettings.FindAll(x => x.InstrumentType.StartsWith("LXPlus")).Select(x => x.InstrumentType));
                return type;
            }

            if (pinName.EndsWith("_SRC"))
            {
                type.AddRange(instrumentTypeSettings.FindAll(x => x.InstrumentType.StartsWith("UltraPac80")).Select(x => x.InstrumentType));
                return type;
            }

            var pinTypeInChannel = _multiChannelMap.SelectMany(p => p.Rows.FindAll(x => x.DeviceUnderTestPinName.Contains(pinName, StringComparison.OrdinalIgnoreCase))).ToList();
            if (pinTypeInChannel.Count != 0)
            {
                return GetInstrumentTypeFromChannel(measPin, instrumentTypeSettings, pinTypeInChannel);
            }

            return GetInstrumentTypeWithoutChannel(instrumentTypeSettings);
        }

        private static List<string> GetInstrumentTypeFromChannel(MeasPin measPin, List<InstrumentTypeSetting> instrumentTypeSettings, List<ChannelMapRow> channelMapRows)
        {
            var type = new List<string>();
            var pinTypesList = new List<string>();
            foreach (InstrumentTypeSetting typePinType in instrumentTypeSettings)
            {
                string[] tmpPinType = typePinType.PinType.Split([','], StringSplitOptions.RemoveEmptyEntries);
                pinTypesList.AddRange(tmpPinType);
            }
            if (pinTypesList.Count == 0)
            {
                pinTypesList.Add("Error Pin Type");
            }

            // For issue #53 if find UVI80 then use it priority, otherwise use UP1600 or other.
            ChannelMapRow? preferred = channelMapRows.FirstOrDefault(
                p => p.DeviceUnderTestPinName?.IndexOf("UVI80", StringComparison.OrdinalIgnoreCase) >= 0
            );

            ChannelMapRow? targetPin = preferred ?? channelMapRows.FirstOrDefault(
                p => ContainsPinType(pinTypesList, p.Type)
                );

            if (targetPin != null)
            {
                AddInstrumentTypeForTargetPin(measPin, instrumentTypeSettings, targetPin, type);
            }
            else if (instrumentTypeSettings.Count == 1)
            {
                type.Add(instrumentTypeSettings.First().InstrumentType);
            }
            else
            {
                type.Add("Error Pin Type");
            }
            return type;
        }

        private static void AddInstrumentTypeForTargetPin(MeasPin measPin, List<InstrumentTypeSetting> instrumentTypeSettings, ChannelMapRow channelMapRow, List<string> type)
        {
            if (channelMapRow.Type.IndexOf("merged", StringComparison.OrdinalIgnoreCase) > 0)
            {
                channelMapRow.Type = channelMapRow.Type[..channelMapRow.Type.IndexOf("merged", StringComparison.OrdinalIgnoreCase)];
            }

            switch (channelMapRow.Type)
            {
                case "GigaDigPos":
                case "GigaDigNeg":
                    type.Add(Gigadig);
                    break;

                case "UltraSource":
                    type.Add(UltraPac80Source);
                    break;
                case "UltraCapture":
                    type.Add(UltraPac80Capture);
                    break;
                case "DCDiffMeter":
                    type.Add(UVI80Vdm);
                    break;

                case "MWSource":
                case "MW":
                    IEnumerable<string> temp = instrumentTypeSettings.Where(y => y.PinType.Split(',').ToList().Exists(p => p.EqualsIgnoreCase(channelMapRow.Type))).Select(p => p.InstrumentType);
                    type.AddRange(temp);
                    break;

                case "DCVS":
                    type.Add(UVS256Meter);
                    break;

                case "DCVI":
                    if (measPin.MeasType.EqualsIgnoreCase("MeasI") || measPin.MeasType.EqualsIgnoreCase("MeasV"))
                    {
                        type.Add(UVI80Meter);
                    }

                    break;

                case "I/O":
                    type.Add(UP1600);
                    break;

                case "Error Frequency":
                    type.Add(ErrorFrequency);
                    break;
                case "Error Pin Type":
                    type.Add(ErrorPinType);
                    break;

            }
        }

        private static List<string> GetInstrumentTypeWithoutChannel(List<InstrumentTypeSetting> instrumentTypeSettings)
        {
            var type = new List<string>();
            if (instrumentTypeSettings.Exists(p => p.InstrumentType.Contains("LitePoint")))
            {
                type.Add(instrumentTypeSettings.First().InstrumentType);
            }
            else if (instrumentTypeSettings.Count > 1)
            {
                List<InstrumentTypeSetting> tmp = instrumentTypeSettings.FindAll(p => p.PinCheck != "V");

                if (tmp.Count != 0)
                {
                    type.Add(tmp.First().InstrumentType);
                }
            }
            else if (instrumentTypeSettings.Count == 1)
            {
                type.Add(instrumentTypeSettings[0].InstrumentType);
            }

            return type;
        }

        private static bool ContainsPinType(List<string> typeList, string pinType)
        {
            if (pinType.IndexOf("merged", StringComparison.OrdinalIgnoreCase) > 0)
            {
                pinType = pinType[..pinType.IndexOf("merged", StringComparison.OrdinalIgnoreCase)];
            }

            bool result = false;
            typeList.ForEach(p =>
            {
                if (Regex.IsMatch(p, pinType))
                {
                    result = true;
                }

            });
            return result;
        }

        public static List<InstrumentTypeSetting> CheckTypeByFreqNew(string measType, Dictionary<string, string> dicPara, List<InstrumentTypeSetting> instrumentTypeSettings)
        {
            if (!dicPara.TryGetValue("freq", out string? value))
            {
                List<InstrumentTypeSetting> typeList = instrumentTypeSettings.FindAll(x => x.HighLimit.Equals(0.0));
                return typeList;
            }
            var freqs = new List<string>();
            foreach (string freq in value.Split(','))
            {
                if (!MyRegex1().IsMatch(freq))
                {
                    freqs.Add(freq + "hz");
                }
                else
                {
                    freqs.Add(freq);
                }
            }
            dicPara["freq"] = string.Join(",", freqs);

            List<string> freqList = [.. dicPara["freq"].Split(',')];
            foreach (string freq in freqList)
            {
                double outputValue =
                    freq.Contains('&') ?
                    freq.Split('&').Max(InstrumentTypeUtility.ConvertToFreqValue) :
                    InstrumentTypeUtility.ConvertToFreqValue(freq);

                List<InstrumentTypeSetting> freqTypeList = instrumentTypeSettings.FindAll(x => !x.HighLimit.Equals(0.0));

                var typeList = freqTypeList.Where(p => outputValue >= p.LowLimit && outputValue <= p.HighLimit).ToList();
                if (typeList.Count == 0)
                {
                    typeList.Add(new InstrumentTypeSetting
                    {
                        InstrumentType = "Error Frequency",
                        PinCheck = "",
                        PinExtraName = "",
                        HighLimit = 0.0,
                        LowLimit = 0.0,
                        Path = "",
                        PinType = ""
                    });
                }
                switch (measType.ToLower())
                {
                    case "wisrc":
                        typeList.RemoveAll(x => x.InstrumentType.Contains("Analyzer") || x.InstrumentType.Contains("Receiver") || x.InstrumentType.Contains("Capture"));
                        break;
                    case "wimeas":
                        typeList.RemoveAll(x => x.InstrumentType.Contains("Source"));
                        break;

                    case "bbsrc":
                        break;
                    case "bbmeas":
                        break;
                }
                return typeList;
            }
            return null!;
        }
    }
}
