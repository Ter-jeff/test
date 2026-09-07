using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using TagDiff.Core.Common;

using TagDiffCore.Utility;

namespace TagDiff.Core.Comparer
{
    public class CharacterizationSheetComparer : IgxlSheetComparerBase
    {
        public Dictionary<int, int> DicVersion25To26HeaderMapping = new()
        {
            {0,0},{1,1},{2,2},{3,3},{4,4},{5,5},{6,6},{7,7},{8,8},{9,9},{10,10},{11,11}, {12,12},{13,13},{14,14},{15,15},{16,16},{17,17},{18,18},{19,19},{20,20},{21,21},{22,22},{23,23},{24,24},{25,25},{26,26},{27,28},{28,29},{29,30},{30,31},{31,32},{32,33}, {33,34},{34,35},{35,36},{36,37},{37,38},{38,39},{39,40},{40,41},{41,42},{42,43},{43,44},{44,46},{45,47},{46,48},{47,49},{48,50},{49,51}
        };

        private readonly Dictionary<int, int> _reverseVersion25To26Mapping;

        public CharacterizationSheetComparer(string sheetName, string baseFile, string compFile, List<SubFlowSheet> subFlowSheets) : base(sheetName, baseFile, compFile, "Setup Name", [])
        {
            UseList = new HashSet<string>([.. subFlowSheets.SelectMany(x => x.Rows).Where(x => x.Opcode.EqualsIgnoreCase("Test")).Where(x => x.Parameter.Split(' ').Length == 2).Select(x => x.Parameter.Split(' ')[1])]);
            KeyColHeaders = ["Setup Name"];
            SheetType = EnumSheetType.DTCharacterizationSheet;
            _reverseVersion25To26Mapping = DicVersion25To26HeaderMapping.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);
            BaseSheetVersion = SheetProvider.GetSheetVersion(BaseSheet);
            if (CompSheet.Count != 0)
            {
                CompareSheetVersion = SheetProvider.GetSheetVersion(CompSheet);
            }
        }

        protected override bool CompareSheetRow(Location baseLocation, Location compLocation, HashSet<int>? skips = null)
        {
            int compareItemRow = compLocation.Row;
            int row = baseLocation.Row;
            return !CompareRow(new Location(row, BaseSheetStartCol), new Location(compareItemRow, compLocation.Col));
        }

        private bool CompareRow(Location baseLocation, Location compLocation)
        {
            bool diff = false;
            int compStartCol = compLocation.Col;
            int baseOffset, compareOffset = -1;
            for (baseOffset = 0; baseOffset <= BaseSheetEndCol - BaseSheetStartCol; baseOffset++)
            {
                if (BaseSheetVersion == CompareSheetVersion)
                {
                    compareOffset = baseOffset;
                }
                else if (BaseSheetVersion == "2.5" && CompareSheetVersion == "2.6")
                {
                    if (!DicVersion25To26HeaderMapping.TryGetValue(baseOffset, out int value))
                    {
                        continue;
                    }

                    compareOffset = value;
                }
                else if (BaseSheetVersion == "2.6" && CompareSheetVersion == "2.5")
                {
                    if (!_reverseVersion25To26Mapping.TryGetValue(baseOffset, out int reverseOffset))
                    {
                        continue;
                    }

                    compareOffset = reverseOffset;
                }

                string baseContext = GetCellSafe(BaseSheet, baseLocation.Row, baseLocation.Col + baseOffset);
                string comContext = GetCellSafe(CompSheet, compLocation.Row, compLocation.Col + compareOffset);
                if (compStartCol > CompSheetEndCol || !baseContext.Trim('=').EqualsIgnoreCase(comContext.Trim('=')))
                {
                    diff = true;
                    SetCellSafe(BaseSheet, baseLocation.Row, baseLocation.Col + baseOffset, $"{baseContext} => {comContext}");
                }
            }
            return diff;
        }
    }
}
