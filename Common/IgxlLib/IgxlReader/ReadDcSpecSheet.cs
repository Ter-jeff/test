using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace IgxlLib.IgxlReader
{
    public class ReadDcSpecSheet : SpecSheetReader<DcSpecSheet, DcSpec>
    {
        public override EnumSheetType SupportedType => EnumSheetType.DTDCSpecSheet;
    }
}
