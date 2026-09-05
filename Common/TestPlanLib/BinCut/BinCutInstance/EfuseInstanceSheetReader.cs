using CommonLib.ErrorReport.ErrorCodes;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class EfuseInstanceSheetReader : BinCutInstanceSheetReader
    {
        public EfuseInstanceSheetReader() : base()
        {
            Type = EnumInstanceSheetType.Efuse;
        }

        public override BinCutInstanceSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            GetDimensions();

            GetFirstHeaderPosition(ConHeaderFlowName);

            GetHeaders();

            BinCutInstanceSheet sheet = ReadSheet(sheetName);

            ClassificationPattern(ref sheet);

            sheet = SplitByCode(sheet);

            CheckPattern(ref sheet, EFuseErrorType.E_MissingPattern_01, EFuseErrorType.E_MissingPattern_01, EFuseErrorType.E_MissingPattern_01);

            return sheet;
        }
    }
}
