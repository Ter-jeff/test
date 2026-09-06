using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class PortMapSheetComparer : IgxlSheetComparerBase
    {
        public PortMapSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Port Name", [])
        {
            KeyColHeaders = ["Port Name", "Pin"];
            SheetType = EnumSheetType.DTPortMapSheet;
        }
    }
}
