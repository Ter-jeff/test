using System;
using System.Collections.Generic;

using IgxlLib.Enums;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

namespace IgxlLib
{
    internal static class IgxlSheetReaderRegistrySpec
    {
        internal static void AddTo(Dictionary<EnumSheetType, (IIgxlSheetReader Reader, Action<object, string> Assign)> mappings, IgxlLoader igxlLoader)
        {
            mappings.Add(EnumSheetType.DTGlobalSpecSheet, (new ReadGlobalSpecSheet(), (data, name) => igxlLoader.GlobalSpecSheet = (GlobalSpecSheet)data));
            mappings.Add(EnumSheetType.DTLevelSheet, (new ReadLevelSheet(), (data, name) => igxlLoader.LevelSheets.Add((LevelSheet)data)));
            mappings.Add(EnumSheetType.DTTimesetBasicSheet, (new ReadTimeSetSheet(), (data, name) => igxlLoader.TimeSetBasicSheets.Add((TimeSetBasicSheet)data)));
            mappings.Add(EnumSheetType.DTACSpecSheet, (new ReadAcSpecSheet(), (data, name) => igxlLoader.AcSpecSheets.Add((AcSpecSheet)data)));
            mappings.Add(EnumSheetType.DTDCSpecSheet, (new ReadDcSpecSheet(), (data, name) => igxlLoader.DcSpecSheets.Add((DcSpecSheet)data)));
        }
    }
}
