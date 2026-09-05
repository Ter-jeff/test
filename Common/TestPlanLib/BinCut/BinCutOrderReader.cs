using System.Collections.Generic;
using System.Linq;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

using TestPlanLib.BinCut.Flow;

namespace TestPlanLib.BinCut
{
    public class BinCutOrderReader(BinCutFlowTables binCutFlowTables) : MySheetReader<BinCutOrderSheet>
    {
        private const string ConHeaderBincut = "Bin Cut";
        private const string ConHeaderPerformanceMode = "Performance Mode";
        private const string ConHeaderOrder = "Order";
        private const string ConHeaderTd = "TD";
        private const string ConHeaderBist = "BIST";
        private const string ConHeaderFunc = "FUNC";

        private readonly BinCutFlowTables _binCutFlowTables = binCutFlowTables;

        private int _indexBincut = -1;
        private int _indexPerformanceMode = -1;
        private int _indexOrder = -1;
        private int _indexTd = -1;
        private int _indexBist = -1;
        private int _indexFunc = -1;

        public override BinCutOrderSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            string sheetName = excelWorksheet.Name;

            var binCutOrderSheet = new BinCutOrderSheet(sheetName);

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                binCutOrderSheet.AddDimensionError();
                return binCutOrderSheet;
            }

            if (!GetFirstHeaderPosition())
            {
                binCutOrderSheet.AddFirstHeaderError(ConHeaderBincut);
                return binCutOrderSheet;
            }

            GetHeaderIndex();

            binCutOrderSheet = ReadSheet(sheetName);

            binCutOrderSheet.Check();

            binCutOrderSheet.CheckPerformanceMode(_binCutFlowTables);

            binCutOrderSheet.Rows = [.. binCutOrderSheet.Rows.OrderBy(x => x.Order)];

            return binCutOrderSheet;
        }

        private BinCutOrderSheet ReadSheet(string sheetName)
        {
            var binCutOrderSheet = new BinCutOrderSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new BinCutOrderRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexBincut != -1)
                {
                    row.Bincut = ExcelWorksheet.GetCellValue(i, _indexBincut).Trim();
                }

                if (_indexPerformanceMode != -1)
                {
                    row.PerformanceMode = ExcelWorksheet.GetCellValue(i, _indexPerformanceMode).Trim();
                }

                if (_indexOrder != -1)
                {
                    _ = int.TryParse(ExcelWorksheet.GetCellValue(i, _indexOrder).Trim(), out int value);
                    row.Order = value;
                }
                if (_indexTd != -1)
                {
                    if (int.TryParse(ExcelWorksheet.GetCellValue(i, _indexTd).Trim(), out int value))
                    {
                        row.Td = value;
                    }
                }
                if (_indexBist != -1)
                {
                    if (int.TryParse(ExcelWorksheet.GetCellValue(i, _indexBist).Trim(), out int value))
                    {
                        row.Bist = value;
                    }
                }
                if (_indexFunc != -1)
                {
                    if (int.TryParse(ExcelWorksheet.GetCellValue(i, _indexFunc).Trim(), out int value))
                    {
                        row.Func = value;
                    }
                }

                binCutOrderSheet.Rows.Add(row);
            }

            binCutOrderSheet.IndexBinCut = _indexBincut;
            binCutOrderSheet.IndexPerformanceMode = _indexPerformanceMode;
            binCutOrderSheet.IndexOrder = _indexOrder;
            binCutOrderSheet.IndexTdx = _indexTd;
            binCutOrderSheet.IndexBist = _indexBist;
            binCutOrderSheet.IndexFunc = _indexFunc;

            return binCutOrderSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string lStrHeader = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (lStrHeader.EqualsIgnoreCase(ConHeaderBincut))
                {
                    _indexBincut = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderPerformanceMode))
                {
                    _indexPerformanceMode = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderOrder))
                {
                    _indexOrder = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderTd))
                {
                    _indexTd = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderBist))
                {
                    _indexBist = i;
                    continue;
                }
                if (lStrHeader.EqualsIgnoreCase(ConHeaderFunc))
                {
                    _indexFunc = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderBincut))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class BinCutOrderSheet : MySheet
    {
        #region Properity
        public List<BinCutOrderRow> Rows { get; set; }

        public int IndexBinCut = -1;
        public int IndexPerformanceMode = -1;
        public int IndexOrder = -1;
        public int IndexTdx = -1;
        public int IndexBist = -1;
        public int IndexFunc = -1;
        #endregion

        #region Constructor
        public BinCutOrderSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion

        public void Check()
        {
            CheckBlockInfo();
        }

        public void CheckPerformanceMode(List<BinCutFlowTable> binCutFlowTables)
        {
            var groups = Rows.GroupBy(x => x.PerformanceMode).ToList();
            foreach (IGrouping<string, BinCutOrderRow> group in groups)
            {
                List<string> modes = GetModes(group.First().Bincut, binCutFlowTables);
                foreach (BinCutOrderRow binCutOrderRow in group)
                {
                    if (!modes.Exists(x => x.EqualsIgnoreCase(binCutOrderRow.PerformanceMode)))
                    {
                        string errorMessage = $"The performance mode {binCutOrderRow.PerformanceMode} can not be found in flow sheet!";
                        AddError(BinCutErrorType.E_FormatError_03, SheetName, binCutOrderRow.RowNum, IndexPerformanceMode, $"The performance mode {binCutOrderRow.PerformanceMode} can not be found in flow sheet!", [binCutOrderRow.PerformanceMode]);
                    }
                }
            }
        }

        private void CheckBlockInfo()
        {
            if (IndexTdx == -1 || IndexBist == -1 || IndexFunc == -1)
            {
                return;
            }

            foreach (BinCutOrderRow row in Rows)
            {
                if (row.Td == null || row.Bist == null || row.Func == null)
                {
                    //string errorMessage = "The columns TD, BIST or FUNC can not be null or empty!";
                    AddError(BinCutErrorType.E_FormatError_04, SheetName, row.RowNum, 0, "The columns TD, BIST or FUNC can not be null or empty!");
                }
            }
        }

        public List<string> GetExtraFlow()
        {
            var rows = Rows.OrderBy(x => x.Order).ToList();
            rows = [.. rows.Where(x => x.Bincut.EqualsIgnoreCase("Search"))];
            return BinCutExtraFlowName.GetExtraPerformanceMode([.. rows.Select(x => x.PerformanceMode).Distinct()]);
        }

        private static List<string> GetModes(string binCut, List<BinCutFlowTable> binCutFlowTables)
        {
            var rows = binCutFlowTables.SelectMany(x => x.Rows).ToList();
            if (binCut.EqualsIgnoreCase("Search"))
            {
                return [.. rows.Where(x => x.TableType == EnumBinCutTableType.Lv).Select(x => x.PerformanceMode)];
            }

            if (binCut.EqualsIgnoreCase("HVCC"))
            {
                return [.. rows.Where(x => x.TableType == EnumBinCutTableType.Hv && x.TableBinType == EnumBinCutTableBinType.Bin1).Select(x => x.PerformanceMode)];
            }

            if (binCut.EqualsIgnoreCase("HVCC - BinX"))
            {
                return [.. rows.Where(x => x.TableType == EnumBinCutTableType.Hv && x.TableBinType == EnumBinCutTableBinType.BinX).Select(x => x.PerformanceMode)];
            }

            if (binCut.EqualsIgnoreCase("HVCC - BinY"))
            {
                return [.. rows.Where(x => x.TableType == EnumBinCutTableType.Hv && x.TableBinType == EnumBinCutTableBinType.BinY).Select(x => x.PerformanceMode)];
            }

            return binCut.EqualsIgnoreCase("Post")
                ? [.. rows.Where(x => x.TableType == EnumBinCutTableType.Post).Select(x => x.PerformanceMode)]
                : [.. rows.Where(x => x.TableType == EnumBinCutTableType.Hv && x.TableBinType == EnumBinCutTableBinType.Bin1).Select(x => x.PerformanceMode)];
        }
    }

    public class BinCutOrderRow : MyRow
    {
        public string Bincut { get; set; }
        public string PerformanceMode { get; set; }
        public int? Order { get; set; }
        public int? Td { get; set; }
        public int? Bist { get; set; }
        public int? Func { get; set; }

        #region Constructor
        public BinCutOrderRow(string sheetName = "")
        {
            SheetName = sheetName;
            Bincut = "";
            PerformanceMode = "";
            Order = 0;
        }
        #endregion
    }
}
