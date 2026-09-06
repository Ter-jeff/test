using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.Basic
{
    public class SubprogramMappingReader : MySheetReader<SubprogramMappingSheet>
    {
        private const string ConSubprogram = "Subprogram";
        private readonly SubprogramMappingSheet _subprogramMappingSheet = new();
        public override SubprogramMappingSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet = excelWorksheet;

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            ReadSheet();

            return _subprogramMappingSheet;
        }

        private void ReadSheet()
        {
            for (int i = StartCol + 1; i <= EndCol; i++)
            {
                if (!string.IsNullOrEmpty(ExcelWorksheet.GetCellValue(StartRow, i).Trim()))
                {
                    SubprogramSetting subprogramSetting = new SubprogramSetting(ExcelWorksheet.GetCellValue(StartRow, i).Trim(), []);
                    for (int j = StartRow + 1; j <= EndRow; j++)
                    {
                        string subFlowName = ExcelWorksheet.GetCellValue(j, StartCol).Trim();
                        bool enable = ExcelWorksheet.GetCellValue(j, i).Trim().EqualsIgnoreCase("TRUE");
                        if (!string.IsNullOrEmpty(subFlowName) && enable)
                        {
                            subprogramSetting.EnableSubFlows.Add(subFlowName);
                        }
                    }
                    if (subprogramSetting.EnableSubFlows.Count != 0)
                    {
                        _subprogramMappingSheet.SubprogramSettings.Add(subprogramSetting);
                    }
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
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConSubprogram))
                    {
                        StartRow = i;
                        StartCol = j;
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
