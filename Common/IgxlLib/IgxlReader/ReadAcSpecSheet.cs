using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace IgxlLib.IgxlReader
{
    public class ReadAcSpecSheet : SpecSheetReader<AcSpecSheet, AcSpec>
    {
        public override EnumSheetType SupportedType => EnumSheetType.DTACSpecSheet;
    }
}
