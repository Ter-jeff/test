using System.Linq;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public class RevisionReader : MySheetReader<bool>
    {
        private const string ConHeaderNotes = "Notes Tab";

        private int _indexNotes = -1;

        public override bool ReadSheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;

            GetDimensions();

            if (!GetFirstHeaderPosition())
            {
                return false;
            }

            return CheckIsShadow();

        }

        private bool CheckIsShadow()
        {
            bool result = false;
            for (int i = StartRow + 1; i <= EndRow; i++)
            {
                string value = ExcelWorksheet.GetCellValue(i, _indexNotes).Trim();
                if (value.Contains("binning_type"))
                {
                    string type = value.Split('=').Last().Trim();
                    if (type == "SHADOW")
                    {
                        result = true;
                    }
                }
            }
            return result;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = EndRow > 10 ? 10 : EndRow;
            int colNum = EndCol > 10 ? 10 : EndCol;
            for (int i = 1; i <= rowNum; i++)
            {
                for (int j = 1; j <= colNum; j++)
                {
                    if (ExcelWorksheet.GetCellValue(i, j).Trim().EqualsIgnoreCase(ConHeaderNotes))
                    {
                        StartRow = i;
                        _indexNotes = j;
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
