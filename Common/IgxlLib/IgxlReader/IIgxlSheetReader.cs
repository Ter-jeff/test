using System.IO;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace IgxlLib.IgxlReader
{
    public interface IIgxlSheetReader
    {
        EnumSheetType SupportedType { get; }

        object Read(object source, string sheetName);
    }

    public interface IIgxlSheetReader<out TSheet> : IIgxlSheetReader where TSheet : IIgxlSheet
    {
        TSheet ReadSheet(ExcelWorksheet excelWorksheet);

        TSheet ReadSheet(Stream stream, string sheetName);
    }
}
