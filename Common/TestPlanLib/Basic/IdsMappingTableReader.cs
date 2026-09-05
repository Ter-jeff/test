using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Basic
{
    public class IdsMappingTableReader : MySheetReader<IdsMappingSheet>
    {
        public const string ConHeaderStage = "Stage";
        public const string ConHeaderSubBlock = "SubBlock";
        public const string ConHeaderPinname = "PinName";
        public const string ConHeaderEfusefieldname = "eFuseFieldName";

        private int _indexStage = -1;
        private int _indexSubBlock = -1;
        private int _indexPinname = -1;
        private int _indexEfusefieldname = -1;

        public override IdsMappingSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            string sheetName = excelWorksheet.Name;

            var iDsMappingSheet = new IdsMappingSheet(sheetName);

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                iDsMappingSheet.AddDimensionError();
                return iDsMappingSheet;
            }

            if (!GetFirstHeaderPosition())
            {
                iDsMappingSheet.AddFirstHeaderError(ConHeaderStage);
                return iDsMappingSheet;
            }

            GetHeaderIndex();

            return ReadSheet(sheetName);
        }

        private IdsMappingSheet ReadSheet(string sheetName)
        {
            var iDsMappingSheet = new IdsMappingSheet(sheetName);
            string currentStage = "";
            string currentSubBlock = "";
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new IdsMappingRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexStage != -1)
                {
                    string content = ExcelWorksheet.GetCellValue(i, _indexStage).Trim();
                    if (!string.IsNullOrEmpty(content))
                    {
                        currentStage = content;
                    }
                }
                if (_indexSubBlock != -1)
                {
                    string content = ExcelWorksheet.GetCellValue(i, _indexSubBlock).Trim();
                    if (!string.IsNullOrEmpty(content))
                    {
                        currentSubBlock = content;
                    }
                }
                if (_indexPinname != -1)
                {
                    row.Pinname = ExcelWorksheet.GetCellValue(i, _indexPinname).Trim();
                }

                if (_indexEfusefieldname != -1)
                {
                    row.Efusefieldname = ExcelWorksheet.GetCellValue(i, _indexEfusefieldname).Trim();
                }

                row.Stage = currentStage;
                row.SubBlock = currentSubBlock;
                iDsMappingSheet.Rows.Add(row);
            }

            iDsMappingSheet.IndexStage = _indexStage;
            iDsMappingSheet.IndexSubBlock = _indexSubBlock;
            iDsMappingSheet.IndexPinname = _indexPinname;
            iDsMappingSheet.IndexEfusefieldname = _indexEfusefieldname;

            return iDsMappingSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderStage))
                {
                    _indexStage = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderSubBlock))
                {
                    _indexSubBlock = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderPinname))
                {
                    _indexPinname = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderEfusefieldname))
                {
                    _indexEfusefieldname = i;
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
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderStage))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }

    public class IdsMappingSheet : MySheet
    {
        public List<IdsMappingRow> Rows { set; get; }
        public Dictionary<string, Dictionary<string, string>> MappingData { set; get; }

        public int IndexStage = -1;
        public int IndexSubBlock = -1;
        public int IndexPinname = -1;
        public int IndexEfusefieldname = -1;

        public IdsMappingSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
            MappingData = [];
        }
    }

    public class IdsMappingRow : MyRow
    {
        public string Stage { set; get; } = "";
        public string SubBlock { get; set; } = string.Empty;
        public string Pinname { set; get; } = "";
        public string Efusefieldname { set; get; } = "";
        public string InstanceName { set; get; } = "";

        #region Constructor
        public IdsMappingRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
