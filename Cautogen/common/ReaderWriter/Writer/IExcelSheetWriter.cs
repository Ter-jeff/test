using OfficeOpenXml;

namespace Cautogen.common.ReaderWriter.Writer
{
    public interface IExcelSheetWriter
    {
        void Write(ExcelWorkbook wb);
    }
}
