using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public class PowerBinningHarvSheetReaderCaicos(List<string> sheetList) : PowerBinningHarvSheetReaderBase<PowerBinningHarvSheetCaicos, PowerBinningHarvRowCaicos>
    {
        private const string InputKeyWord = "INPUT:";
        private const string Powerbinning = "PASS: power_binning";
        private const string FuseName2 = "PASS: fuse_name2";
        private const string FailHeader = "FAIL";
        private const string ConHeaderComment = "Comment";

        private readonly List<string> _sheetList = sheetList;

        protected override bool GetFirstHeaderList()
        {
            for (int j = StartCol; j <= EndCol; j++)
            {
                string text = ExcelWorksheet.GetCellValue(1, j);
                if (text.EqualsIgnoreCase(Powerbinning))
                {
                    IndexPowerbinning = j;
                    continue;
                }
                if (text.EqualsIgnoreCase(FuseName2))
                {
                    IndexFuseName2Index = j;
                    continue;
                }
                if (text.EqualsIgnoreCase(FailHeader))
                {
                    IndexFailCommand = j;
                    continue;
                }
                if (text.EqualsIgnoreCase(ConHeaderComment))
                {
                    IndexCommentIndex = j;
                    continue;
                }
                if (text.Trim().StartsWithIgnoreCase(InputKeyWord))
                {
                    InputHeaderList.Add(text.Replace(InputKeyWord, "").Trim().ToUpper(), j);
                    continue;
                }

                foreach (string sheet in _sheetList)
                {
                    string sheetName = Path.GetFileNameWithoutExtension(sheet);
                    if (text.EqualsIgnoreCase(sheetName))
                    {
                        PowerBinningSheetHeaderList.Add(text, j);
                    }
                }
            }
            return true;
        }
    }

    public class PowerBinningHarvSheetCaicos : PowerBinningHavrSheet
    {
        public new List<PowerBinningHarvRowCaicos> Rows { get; }

        public PowerBinningHarvSheetCaicos(string sheetname) : base(sheetname)
        {
            SheetName = sheetname;
            Rows = [];
        }
    }

    public class PowerBinningHarvRowCaicos : PowerBinningHarvRow
    {
        public PowerBinningHarvRowCaicos(string sheetName = "")
        {
            SheetName = sheetName;
            Inputinfo = [];
            SheetInfo = [];
            PowerBinning = "";
            FuseName2 = "";
            FailCommand = "";
            Comment = "";
        }
    }
}
