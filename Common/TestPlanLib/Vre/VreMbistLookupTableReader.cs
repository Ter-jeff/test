using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;


namespace TestPlanLib
{
    public class VreMbistLookupTableReader : MySheetReader<VreMbistLookupTable>
    {
        private const string ConHarvesting = "Harvesting";
        private const string ConServer = "Server";
        private const string ConMemeryGroup = "Memory Group";
        private const string ConPmode = "P mode";
        private const string ConExcludePattern = "Exclude Pattern";


        private int _indexHarvesting = -1;
        private int _indexServer = -1;
        private int _indexMermoryGroup = -1;
        private int _indexPmode = -1;
        private int _indexExcludePattern = -1;
        private Dictionary<string, int> _headerIndex = new Dictionary<string, int>();
        private readonly VreMbistLookupTable _oreMbistLookupTableTable = new VreMbistLookupTable();
        public override VreMbistLookupTable? ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            ExcelWorksheet = worksheet;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            ReadSheet();

            _oreMbistLookupTableTable.SheetName = ExcelWorksheet.Name;
            return _oreMbistLookupTableTable;
        }

        private void ReadSheet()
        {
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new OreMbistLookupRow();
                if (_indexHarvesting != -1)
                {
                    row.Harvesting = ExcelWorksheet.GetCellValue(i, _indexHarvesting).Trim();
                }
                if (_indexServer != -1)
                {
                    row.Server = ExcelWorksheet.GetCellValue(i, _indexServer).Trim();
                }
                if (_indexMermoryGroup != -1)
                {
                    row.MemoryGroup = ExcelWorksheet.GetCellValue(i, _indexMermoryGroup).Trim();
                }
                if (_indexPmode != -1)
                {
                    row.Pmode = ExcelWorksheet.GetCellValue(i, _indexPmode).Trim();
                }
                if (_indexExcludePattern != -1)
                {
                    row.ExcludePattern = ExcelWorksheet.GetCellValue(i, _indexExcludePattern).Trim();
                }
                _oreMbistLookupTableTable.Rows.Add(row);
            }
            _oreMbistLookupTableTable.HeaderIndex = _headerIndex;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                bool foundIndex = true;
                if (header.EqualsIgnoreCase(ConHarvesting))
                {
                    _indexHarvesting = i;
                }
                else if (header.EqualsIgnoreCase(ConServer))
                {
                    _indexServer = i;
                }
                else if (header.EqualsIgnoreCase(ConMemeryGroup))
                {
                    _indexMermoryGroup = i;
                }
                else if (header.EqualsIgnoreCase(ConPmode))
                {
                    _indexPmode = i;
                }
                else if (header.EqualsIgnoreCase(ConExcludePattern))
                {
                    _indexExcludePattern = i;
                }
                else
                {
                    foundIndex = false;
                }
                if (foundIndex)
                {
                    _headerIndex.Add(header, i);
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow;
            int colNum = EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHarvesting))
                    {
                        StartRow = i;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
