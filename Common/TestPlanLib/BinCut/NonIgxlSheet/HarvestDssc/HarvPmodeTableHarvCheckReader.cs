using System.Collections.Generic;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc
{
    public class HarvPmodeTableHarvCheckReader : MySheetReader<HarvPmodeTableHarvCheckSheet>
    {
        private const string ConHeaderHarvCheck = "HARV_Check";
        private const string ConHeaderDeviceCondition = "Device Condition";
        private const string ConHeaderDisableCore = "Disable_Core";

        private int _indexHarvCheck = -1;
        private int _indexDeviceCondition = -1;
        private int _indexDisableCore = -1;

        public override HarvPmodeTableHarvCheckSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            string sheetName = excelWorksheet.Name;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            GetHeaderIndex();

            HarvPmodeTableHarvCheckSheet harvPmodeTableHarvCheckSheet = ReadSheet(sheetName);

            return harvPmodeTableHarvCheckSheet;
        }

        private HarvPmodeTableHarvCheckSheet ReadSheet(string sheetName)
        {
            var harvPmodeTableHarvCheckSheet = new HarvPmodeTableHarvCheckSheet(sheetName);
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                var row = new HarvPmodeTableHarvCheckRow(sheetName)
                {
                    RowNum = i
                };
                if (_indexHarvCheck != -1)
                {
                    row.HarvCheck = ExcelWorksheet.GetCellValue(i, _indexHarvCheck).Trim();
                }

                if (_indexDeviceCondition != -1)
                {
                    row.DeviceCondition = ExcelWorksheet.GetCellValue(i, _indexDeviceCondition).Trim();
                }

                if (_indexDisableCore != -1)
                {
                    row.DisableCore = ExcelWorksheet.GetCellValue(i, _indexDisableCore).Trim();
                }

                if (string.IsNullOrEmpty(row.HarvCheck))
                {
                    break;
                }

                harvPmodeTableHarvCheckSheet.Rows.Add(row);
            }

            harvPmodeTableHarvCheckSheet.HarvCount = harvPmodeTableHarvCheckSheet.Rows.FindAll(x => x.HarvCheck != null).Count;
            harvPmodeTableHarvCheckSheet.IndexHarvCheck = _indexHarvCheck;
            harvPmodeTableHarvCheckSheet.IndexDeviceCondition = _indexDeviceCondition;
            harvPmodeTableHarvCheckSheet.IndexDisableCore = _indexDisableCore;

            return harvPmodeTableHarvCheckSheet;
        }

        private void GetHeaderIndex()
        {
            for (int i = StartCol; i <= EndCol; i++)
            {
                string header = ExcelWorksheet.GetCellValue(StartRow, i).Trim();
                if (header.EqualsIgnoreCase(ConHeaderHarvCheck))
                {
                    _indexHarvCheck = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderDeviceCondition))
                {
                    _indexDeviceCondition = i;
                    continue;
                }
                if (header.EqualsIgnoreCase(ConHeaderDisableCore))
                {
                    _indexDisableCore = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            //int rowNum = EndRowNumber > 10 ? 10 : EndRowNumber;
            //int colNum = EndColNumber > 10 ? 10 : EndColNumber;
            for (int i = 1; i <= EndRow; i++)
            //for (int j = 1; j <= colNum; j++)
            {
                if (ExcelWorksheet.GetCellValue(i, 1).Trim().EqualsIgnoreCase(ConHeaderHarvCheck))
                {
                    StartRow = i;
                    return true;
                }
            }
            return false;
        }
    }

    public class HarvPmodeTableHarvCheckSheet : MySheet
    {
        public List<HarvPmodeTableHarvCheckRow> Rows { set; get; }

        public int IndexHarvCheck = -1;
        public int IndexDeviceCondition = -1;
        public int IndexDisableCore = -1;
        public int HarvCount = -1;

        #region Constructor
        public HarvPmodeTableHarvCheckSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion

        public Dictionary<string, string> GetFlags(string groupCondition)
        {
            var flags = new Dictionary<string, string>();
            string[] arr = groupCondition.Split(',');
            foreach (string item in arr)
            {
                if (Rows.Exists(x => x.DisableCore.EqualsIgnoreCase(item)))
                {
                    HarvPmodeTableHarvCheckRow row = Rows.Find(x => x.DisableCore.EqualsIgnoreCase(item))!;
                    flags.Add(row.HarvCheck, row.DeviceCondition);
                }
            }
            return flags;
        }
    }

    public class HarvPmodeTableHarvCheckRow : MyRow
    {
        public string HarvCheck { set; get; } = "";
        public string DeviceCondition { set; get; } = "";
        public string DisableCore { set; get; } = "";

        #region Constructor
        public HarvPmodeTableHarvCheckRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
