using CommonLib.ErrorReport.ErrorCodes;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class NonBinCutInstanceSheetReader : BinCutInstanceSheetReader
    {
        public NonBinCutInstanceSheetReader() : base()
        {
            Type = EnumInstanceSheetType.Scan;
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

            CheckPattern(ref sheet, BinCutErrorType.E_Pattern_05, BinCutErrorType.E_Pattern_06, BinCutErrorType.W_Pattern_06);

            return sheet;
        }
    }
}
