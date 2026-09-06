using CommonLib.ErrorReport.ErrorCodes;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class EvsInstanceSheetReader : BinCutInstanceSheetReader
    {
        public EvsInstanceSheetReader() : base()
        {
            Type = EnumInstanceSheetType.Evs;
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

            CheckPattern(ref sheet, EvsErrorType.E_Pattern_01, EvsErrorType.E_Pattern_01, EvsErrorType.E_Pattern_01);

            return sheet;
        }
    }
}
