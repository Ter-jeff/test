using IgxlLib.Const;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public class MainFlow : SubFlowSheet
    {
        public override string SheetType => "DTFlowtableSheet";
        public override string IgxlSheetName => IgxlSheetNames.FlowTable;

        public MainFlow(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
        }

        public MainFlow(string sheetName, string sourceSheet = "") : base(sheetName)
        {
            SourceInfo.Name = sourceSheet;
        }
    }
}
