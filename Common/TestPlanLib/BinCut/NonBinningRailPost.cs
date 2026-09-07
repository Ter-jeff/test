using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public class NonBinningRailPost : NonBinningRail
    {
        #region Field
        private const string ConHeaderPerfMode = "Performance.*";
        private readonly int _count;
        #endregion

        public NonBinningRailPost(string folder, ExcelWorksheet excelWorksheet, int count)
            : base(folder, excelWorksheet)
        {
            StartHeader = ConHeaderPerfMode;
            _count = count;
        }

        protected override string ReturnPostFileName(bool isCs, string sheetName)
        {
            return sheetName.Contains("post", System.StringComparison.OrdinalIgnoreCase)
                ? (!isCs ? BinCutConst.ConNonBinningRailOutsideFileName : BinCutConst.ConBincutAteConditionOutsideFileName) + "_" + _count
                : !isCs ? BinCutConst.ConNonBinningRailOutsideFileName : BinCutConst.ConBincutAteConditionOutsideFileName;
        }
    }
}
