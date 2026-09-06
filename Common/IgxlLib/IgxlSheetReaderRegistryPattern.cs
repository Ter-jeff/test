using System;
using System.Collections.Generic;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

namespace IgxlLib
{
    internal static class IgxlSheetReaderRegistryPattern
    {
        internal static void AddTo(Dictionary<EnumSheetType, (IIgxlSheetReader Reader, Action<object, string> Assign)> mappings, IgxlLoader igxlLoader)
        {
            mappings.Add(EnumSheetType.DTPatternSubroutineSheet, (new ReadPatSubroutineSheet(), (data, name) => igxlLoader.PatSetSubSheets.Add((PatSetSubSheet)data)));
            void Action(object data, string name)
            {
                var casted = (PatSetSheet)data;
                if (name.EqualsIgnoreCase("PatSets_All"))
                {
                    igxlLoader.PatSetsAll = casted;
                }
                else
                {
                    igxlLoader.PatSetSheets.Add(casted);
                }
            }
            mappings.Add(EnumSheetType.DTPatternSetSheet, (new ReadPatSetSheet(), Action));
        }
    }
}
