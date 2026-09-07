using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.BinCut.NonIgxlSheet;
using TestPlanLib.Singleton;
using TestPlanLib.Static;
using TestPlanLib.Utility;

namespace TestPlanLib.BinCut
{
    public partial class VddBinningDef : BinCutNonIgxlBase
    {
        private const int StartRow = 3;

        private const string BinNumerConditionLvcc = "LVCC";
        private const string BinNumerConditionBining = "Bining";
        private const string ConBinnedHeader = "Binned";
        private const string ConStartHeader = "Domain";

        private const string ConEndHeader = "Comment";
        private const string ConSpi = "RTOS";
        private const string ConFunc = "Func";
        private const string ConSoftbin = "Softbin";
        private const string ConHardBin = "HardBin";
        private const string ConLvccFail = "LVCC Fail";
        private const string ConBinCutListEqual = "Bin Cut List =";
        private const string ConSoftBinEqual = "col_soft_bin =";

        [GeneratedRegex("BinX|BinY", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("_BinX", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex("_BinY", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();

        public string BinCutList = "";
        public int Index;

        public VddBinningDef(string folder, ExcelWorksheet excelWorksheet)
            : base(excelWorksheet, folder)
        {
            StartHeader = ConStartHeader;
        }

        protected override void EditSheet(DataTable dataTable, ref string errMsg, bool isCsharp)
        {
            VddBinningDefHelpers.AdjustStartRow(dataTable);

            RemoveBlankTail(dataTable);

            BinCutNonIgxlBaseHelpers.RemoveColumnsAfterHeader(ConStartHeader, ConEndHeader, dataTable, ref errMsg);
            VddBinningDefHelpers.RemovedBlankBeforeData(dataTable, StartRow);

            if (dataTable.Rows[0][1].ToString()!.Contains('.'))
            {
                dataTable.Rows[0][1] = dataTable.Rows[0][1].ToString()!.Replace('.', 'P');
            }
            else
            {
                dataTable.Rows[0][1] += "P0";
            }

            dataTable.Rows[0][3] = ConBinCutListEqual;

            if (_regex.IsMatch(dataTable.TableName))
            {
                dataTable.Rows[0][4] = Index;
            }
            else
            {
                dataTable.Rows[0][4] = BinCutList;
            }

            dataTable.Rows[0][9] = ConSoftBinEqual;

            dataTable.Rows[0][10] = dataTable.Columns.Count + 1;

            int endColumn = dataTable.Columns.Count;
            VddBinningDefHelpers.AddBaseVoltage(dataTable, BaseVoltage);
            VddBinningDefHelpers.AddStepSize(dataTable, StepSize);

            AddBinNumber(dataTable, endColumn, isCsharp);

            OverwriteBinNumber(Worksheet, dataTable, endColumn);
        }

        private static void OverwriteBinNumber(ExcelWorksheet excelWorksheet, DataTable dataTable, int endColumn)
        {
            ExcelWorksheet? sheetConfig = null;
            if (sheetConfig != null)
            {
                var binNumberAssignmentReader = new BinNumberAssignmentReader();
                BinNumberAssignmentSheet? binNumberAssignmentSheet = binNumberAssignmentReader.ReadSheet(sheetConfig);
                OverWirteBinNumber(dataTable, endColumn, binNumberAssignmentSheet!);
            }
            if (sheetConfig == null)
            {
                ExcelWorksheet? sheet = excelWorksheet.Workbook.Worksheets["Bincut_BinNumberAssignment"];
                if (sheet != null)
                {
                    var binNumberAssignmentReader = new BinNumberAssignmentReader();
                    BinNumberAssignmentSheet? binNumberAssignmentSheet = binNumberAssignmentReader.ReadSheet(sheet);
                    OverWirteBinNumber(dataTable, endColumn, binNumberAssignmentSheet!);
                }
            }
        }

        private static void OverWirteBinNumber(DataTable dataTable, int endColumn, BinNumberAssignmentSheet binNumberAssignmentSheet)
        {
            const int modeIndex = 2;
            foreach (BinNumberAssignmentRow item in binNumberAssignmentSheet.Rows)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    if (dataTable.Rows[i][modeIndex].ToString()!.EqualsIgnoreCase(item.Mode))
                    {
                        dataTable.Rows[i][endColumn + 1] = item.Binningfail1;
                        dataTable.Rows[i][endColumn + 2] = item.Lvccfail1;
                        dataTable.Rows[i][endColumn + 3] = item.Binningfail2;
                        dataTable.Rows[i][endColumn + 4] = item.Lvccfail2;
                        dataTable.Rows[i][endColumn + 5] = item.Binningfail3;
                        dataTable.Rows[i][endColumn + 6] = item.Lvccfail3;
                        dataTable.Rows[i][endColumn + 7] = item.Binningfail4;
                        dataTable.Rows[i][endColumn + 8] = item.Lvccfail4;
                        dataTable.Rows[i][endColumn + 9] = item.Binningfail5;
                        dataTable.Rows[i][endColumn + 10] = item.Lvccfail5;
                        dataTable.Rows[i][endColumn + 11] = item.Binningfail6;
                        dataTable.Rows[i][endColumn + 12] = item.Lvccfail6;
                    }
                }
            }
        }

        protected override string ReturnFileName(bool isCs)
        {
            if (isCs)
            {
                return BinCutConst.ConBincutEqnFileName;
            }

            if (_regex2.IsMatch(SheetName))
            {
                return BinCutConst.ConVddBinningDefFileName2;
            }
            return _regex3.IsMatch(SheetName) ? BinCutConst.ConVddBinningDefFileName3 : BinCutConst.ConVddBinningDefFileName1;
        }

        private void AddBinNumber(DataTable dataTable, int endColumn, bool isCsharp)
        {
            //, "FALSE", "ATE" };
            List<string> allowRow = ["TRUE"];
            const int bintypesIdx = 0;
            const int lvccAndBinningsIdx = 1;
            const int softAndHardBinsIdx = 2;

            VddBinningDefHelpers.FormatByIsCsharp(isCsharp, out List<string> bintypes, out List<string> lvccAndBinnings);

            List<string> softAndHardBins = [ConSoftbin, ConHardBin];
            var columnDics = new Dictionary<string, int>();
            foreach (string bintype in bintypes)
            {
                foreach (string softAndHardBin in softAndHardBins)
                {
                    foreach (string lvccAndBinning in lvccAndBinnings)
                    {
                        columnDics.Add(string.Join("_", new List<string> { bintype, lvccAndBinning, softAndHardBin }), endColumn++);
                    }
                }
            }

            if (!isCsharp)
            {
                for (int i = 0; i < columnDics.Count; i++)
                {
                    dataTable.Columns.Add();
                }

                //Add bintypes
                foreach (KeyValuePair<string, int> columnDic in columnDics)
                {
                    dataTable.Rows[bintypesIdx][columnDic.Value] = columnDic.Key.Split('_')[bintypesIdx];
                }

                // Add lvccAndBinnings
                foreach (KeyValuePair<string, int> columnDic in columnDics)
                {
                    dataTable.Rows[lvccAndBinningsIdx][columnDic.Value] = columnDic.Key.Split('_')[lvccAndBinningsIdx];
                }

                // Add softAndHardBins
                foreach (KeyValuePair<string, int> columnDic in columnDics)
                {
                    dataTable.Rows[softAndHardBinsIdx][columnDic.Value] = columnDic.Key.Split('_')[softAndHardBinsIdx];
                }
            }
            int modeColumn = VddBinningDefHelpers.FindSpecColumn("Mode", dataTable);
            int binned = VddBinningDefHelpers.FindSpecColumn(ConBinnedHeader, dataTable);
            for (int i = StartRow; i < dataTable.Rows.Count; i++)
            {
                string performanceMode = dataTable.Rows[i][modeColumn].ToString()!;
                if (SheetName.EqualsIgnoreCase(NeededSheets.Binning) ||
                    SheetName.EqualsIgnoreCase(NeededSheets.BinningBinX) ||
                    SheetName.EqualsIgnoreCase(NeededSheets.BinningBinY))
                {
                    if (!allowRow.Any(x => x.EqualsIgnoreCase(dataTable.Rows[i][binned].ToString()!)))
                    {
                        continue;
                    }
                }

                //3,4,7,8,11,12 -> Hard bin
                foreach (string lvccAndBinning in lvccAndBinnings)
                {
                    string condition = lvccAndBinning == ConLvccFail ? BinNumerConditionLvcc : BinNumerConditionBining;
                    foreach (KeyValuePair<string, int> column in columnDics.Where(x => x.Key.Split('_')[lvccAndBinningsIdx] == lvccAndBinning))
                    {
                        string category1 = performanceMode[..^1];
                        string category2 = column.Key.Split('_')[bintypesIdx] + "_BV";
                        VddBinningDefHelpers.AddBincutBinNums(dataTable, isCsharp, i, performanceMode, column, category1, category2);
                    }
                }
                foreach (KeyValuePair<string, int> softbinItem in columnDics.Where(x => x.Key.Split('_')[softAndHardBinsIdx] == ConSoftbin))
                {
                    List<string> items = [.. softbinItem.Key.Split('_')];
                    string lvccAndBinning = items[lvccAndBinningsIdx];
                    string bintype = items[bintypesIdx];
                    string binName = items[bintypesIdx].Replace(ConSpi, ConFunc);
                    string idsOrBv = lvccAndBinning == ConLvccFail ? "BV" : "IDS";
                    string key = performanceMode + "," + bintype + "_BV";
                    if (BinNumberSingleton.BincutBinNums.TryGetValue(key, out BinNumber.BinNumResult? value) && !isCsharp)
                    {
                        dataTable.Rows[i][softbinItem.Value] = value.SoftBin.ToString("G15");
                    }
                }
            }
        }
    }
}
