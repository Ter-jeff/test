using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using CommonReaderLib;

using OfficeOpenXml;

namespace TestPlanLib.BinCut
{
    public class PowerBinningHarvSheetReader(List<string> sheetList) : PowerBinningHarvSheetReaderBase<PowerBinningHavrSheet, PowerBinningHarvRow>
    {
        private const string InputKeyWord = "INPUT:";
        private const string FuseName2 = "PASS: fuse_name2";
        private const string FailHeader = "FAIL";
        private const string ConHeaderComment = "Comment";

        private readonly List<string> _sheetList = sheetList;

        protected override bool GetFirstHeaderList()
        {
            for (int j = StartCol; j <= EndCol; j++)
            {
                string text = ExcelWorksheet.GetCellValue(1, j);
                if (text.StartsWith("PASS:") && text.Split([':'], StringSplitOptions.RemoveEmptyEntries).Last().Contains("POWER_BINNING", StringComparison.OrdinalIgnoreCase))
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

    public class PowerBinningHavrSheet : MySheet
    {
        public List<PowerBinningHarvRow> Rows { get; }

        public int IndexPowerbinning = -1;
        public int IndexFuseName2Index = -1;
        public int IndexFailCommand = -1;
        public int IndexCommentIndex = -1;
        public Dictionary<string, int> InputHeader = [];
        #region Constructor
        public PowerBinningHavrSheet(string sheetname)
        {
            SheetName = sheetname;
            Rows = [];
        }
        #endregion
    }

    public class PowerBinningHarvRow : MyRow
    {
        public List<Dictionary<string, string>> Inputinfo;
        public List<Dictionary<string, string>> SheetInfo;
        public string PowerBinning;
        public string FuseName2;
        public string FailCommand;
        public string Comment;

        #region Constructor
        public PowerBinningHarvRow(string sheetName = "")
        {
            SheetName = sheetName;
            Inputinfo = [];
            SheetInfo = [];
            PowerBinning = "";
            FuseName2 = "";
            FailCommand = "";
            Comment = "";
        }
        #endregion
    }
}
