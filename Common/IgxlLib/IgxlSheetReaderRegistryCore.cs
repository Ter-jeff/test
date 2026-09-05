using System;
using System.Collections.Generic;

using IgxlLib.Enums;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

namespace IgxlLib
{
    internal static class IgxlSheetReaderRegistryCore
    {
        internal static void AddTo(Dictionary<EnumSheetType, (IIgxlSheetReader Reader, Action<object, string> Assign)> mappings, IgxlLoader igxlLoader)
        {
            mappings.Add(EnumSheetType.DTFlowtableSheet, (new ReadFlowSheet(), (data, name) => igxlLoader.FlowSheets.Add((SubFlowSheet)data)));
            mappings.Add(EnumSheetType.DTTestInstancesSheet, (new ReadInstanceSheet(), (data, name) => igxlLoader.InstanceSheets.Add((InstanceSheet)data)));
            mappings.Add(EnumSheetType.DTBintablesSheet, (new ReadBinTableSheet(), (data, name) => igxlLoader.BinTableSheets.Add((BinTableSheet)data)));
            mappings.Add(EnumSheetType.DTJobListSheet, (new ReadJobListSheet(), (data, name) => igxlLoader.JobListSheet = (JobListSheet)data));
            mappings.Add(EnumSheetType.DTPinMap, (new ReadPinMapSheet(), (data, name) => igxlLoader.PinMapSheets.Add((PinMapSheet)data)));
            mappings.Add(EnumSheetType.DTChanMap, (new ReadChanMapSheet(), (data, name) => igxlLoader.ChannelMapSheets.Add((ChannelMapSheet)data)));
            mappings.Add(EnumSheetType.DTCharacterizationSheet, (new ReadCharacterizationSheet(), (data, name) => igxlLoader.CharSheets.Add((CharSheet)data)));
        }
    }
}
