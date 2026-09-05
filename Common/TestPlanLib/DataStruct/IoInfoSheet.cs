using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

namespace TestPlanLib.DataStruct
{
    public class IoInfoSheet
    {
        public string SheetName { get; } = "";

        public Dictionary<string, List<IoInfoRow>> BlockIoInfo { get; } = [];

        // mock
        public IoInfoSheet()
        {
        }

        public IoInfoSheet(string sheetName, Dictionary<string, List<IoInfoRow>> blockIoInfo)
        {
            SheetName = sheetName;
            BlockIoInfo = blockIoInfo;
        }

        public List<IoInfoRow> GetBlockIoInfo(string block)
        {
            if (BlockIoInfo.TryGetValue(block, out List<IoInfoRow>? info))
            {
                return info;
            }
            return [];
        }

        public HashSet<string> GetBlockPins(string block = "")
        {
            if (BlockIoInfo.TryGetValue(block, out List<IoInfoRow>? value))
            {
                return value.Select(x => x.PinGrpName).ToHashSet(StringExtensions.IgnoreCase);
            }
            return [];
        }
    }
}
