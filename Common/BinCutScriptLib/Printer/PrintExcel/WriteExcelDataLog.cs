using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Base.Line;

using CommonLib.Extension;

using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

namespace BinCutScriptLib.Printer.PrintExcel
{
    internal class WriteExcelDataLog
    {
        public object MisValue = Type.Missing;
        public ExcelPackage XlPackage;
        public ExcelWorkbook XlWorkBook;
        public string DestExcel = string.Empty;
        public List<SiteInfo> AllDiceInfos = [];
        public List<SiteInfo> AllDiceInfosCp1 = [];
        public List<string> InstNameList = [];
        public Dictionary<int, List<AdjustVddBinningRow>> AdjustVddBinningRows = [];
        private readonly List<Tuple<SiteInfo, List<PowerBinningLineRow>>> _powerBinningLines = [];
        private readonly List<Tuple<int, string>> _binTable = [];

        public WriteExcelDataLog(ExcelPackage excelPackage)
        {
            XlPackage = excelPackage;
            XlWorkBook = excelPackage.Workbook;
        }

        public WriteExcelDataLog(string outPath, List<SiteInfo> siteInfos, Dictionary<int, List<AdjustVddBinningRow>> adjustVddBinningRows,
            List<Tuple<SiteInfo, List<PowerBinningLineRow>>> powerBinningLines, List<Tuple<int, string>> binTable, List<string> instNameList)
        {
            string dirName = Path.GetDirectoryName(outPath)!;
            string fileName = Path.GetFileNameWithoutExtension(outPath);
            DestExcel = Path.Combine(dirName, fileName + ".xlsx");
            AllDiceInfos = [.. siteInfos];
            AllDiceInfosCp1 = siteInfos;
            InstNameList = instNameList;
            AdjustVddBinningRows = adjustVddBinningRows;
            _powerBinningLines = powerBinningLines;
            _binTable = binTable;
            XlPackage = new ExcelPackage();
            XlWorkBook = XlPackage.Workbook;
        }

        public void WritePowerBinningList()
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("PowerBinning");

            //ptring title
            int currRow = 1;
            var titles = new List<string> { "Type", "Parameter" };
            foreach (Tuple<SiteInfo, List<PowerBinningLineRow>> oneDice in _powerBinningLines)
            {
                titles.Add($"X:{oneDice.Item1.XCoor} Y:{oneDice.Item1.YCoor}");
            }

            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(titles.ToArray());

            List<string> allItemName = [.. _powerBinningLines.SelectMany(x => x.Item2).Select(x => x.PowerName).Distinct()];
            int max = allItemName.Count;
            for (int i = 0; i < max; i++)
            {
                string item = allItemName[i];
                List<string> arrayPTotal = ["P_Total", item];
                List<string> data = [];
                foreach (Tuple<SiteInfo, List<PowerBinningLineRow>> oneDice in _powerBinningLines)
                {
                    if (oneDice.Item2.Exists(x => x.PowerName.EqualsIgnoreCase(item)))
                    {
                        PowerBinningLineRow row = oneDice.Item2.Find(x => x.PowerName.EqualsIgnoreCase(item))!;
                        data.Add(row.Value.ToString());
                    }
                    else
                    {
                        data.Add("N/A");
                    }
                }
                arrayPTotal.AddRange(data);
                currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(arrayPTotal.ToArray());
            }

            currRow += 3;
            xlWorkSheet.Cells[currRow, 1].Value = "PowerBinning Check";
            string[] list = [.. AllDiceInfos.Select(x => x.CheckResult.IsPowerBinningPass ? "P" : "F")];
            currRow = xlWorkSheet.Cells[currRow, 3].PrintExcelRow(list);

            xlWorkSheet.Cells[currRow, 1].Value = "IsPowerBinningFail";
            string[] powerBinningFail = [.. AllDiceInfos.Select(x => x.PowerBinningFail)];
            currRow = xlWorkSheet.Cells[currRow, 3].PrintExcelRow(powerBinningFail);

            xlWorkSheet.Cells[currRow, 1].Value = "IsPowerBinningBinXFail";
            string[] powerBinningBinXFail = [.. AllDiceInfos.Select(x => x.PowerBinningBinXFail)];
            currRow = xlWorkSheet.Cells[currRow, 3].PrintExcelRow(powerBinningBinXFail);

            IExcelConditionalFormattingEqual cond1 = xlWorkSheet.ConditionalFormatting.AddEqual(new ExcelAddress("1:" + currRow));
            cond1.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cond1.Style.Fill.BackgroundColor.Color = Color.Red;
            cond1.Style.Font.Color.Color = Color.White;
            cond1.Formula = "\"F\"";

            xlWorkSheet.Cells.TryAutoFitColumns();
        }

        public void WriteSummary()
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("Summary");

            //print title
            int currRow = 1;
            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(["Soft", "Bin", "ECID", "PM", "IDS", "EQN", "VDD", "CP"]);

            //print data
            foreach (KeyValuePair<int, List<AdjustVddBinningRow>> adjustVddBinningRow in AdjustVddBinningRows)
            {
                IEnumerable<IGrouping<string, AdjustVddBinningRow>> groups = adjustVddBinningRow.Value.GroupBy(x => x.PowerName + "_" + x.XCoor + "_" + x.YCoor);
                foreach (IGrouping<string, AdjustVddBinningRow> group in groups)
                {
                    var rows = group.ToList();
                    int xCoor = group.First().XCoor;
                    int yCoor = group.First().YCoor;
                    Tuple<SiteInfo, List<PowerBinningLineRow>>? die = _powerBinningLines.Find(x => x.Item1.XCoor == xCoor && x.Item1.YCoor == yCoor);
                    int soft = die == null ? -1 : die.Item1.Sort;
                    int bin = die == null ? -1 : die.Item1.Bin;
                    string ecid = $"X:{xCoor} Y:{yCoor}";
                    string pm = group.First().PowerName;
                    double ids = GetValue(rows, "IDS");
                    double eq = GetValue(rows, "EQN");
                    double vdd = GetValue(rows, "VDD");
                    double cp = GetValue(rows, "CP");
                    List<object> array =
                    [
                        soft,
                        bin,
                        ecid,
                        pm,
                        ids,
                        eq,
                        vdd,
                        cp,
                    ];
                    xlWorkSheet.Cells[currRow++, 1].PrintExcelRow(array.ToArray());
                }

            }

            xlWorkSheet.Cells.TryAutoFitColumns();
        }

        private static double GetValue(List<AdjustVddBinningRow> adjustVddBinningRows, string name)
        {
            AdjustVddBinningRow? row = adjustVddBinningRows.Find(x => x.Type.EqualsIgnoreCase(name));
            if (row == null)
            {
                return -1;
            }

            return row.Value;
        }

        public void WriteHistogram()
        {
            ExcelWorksheet xlWorkSheet = XlWorkBook.AddSheet("Histogram");

            Dictionary<string, List<string>> powerNamesVsBinEq = GetpowerNamesVsBinEq();

            //print title
            int currRow = 1;
            List<string> title = ["Mode", "EQ NO.", "Total"];
            int siteCount = AllDiceInfos.Count;
            for (int site = 0; site < siteCount; site++)
            {
                title.Add("Site " + site);
            }

            currRow = xlWorkSheet.Cells[currRow, 1].PrintExcelRow(title.ToArray());

            //print data
            List<List<string>> data = [];
            foreach (KeyValuePair<string, List<string>> dic in powerNamesVsBinEq)
            {
                foreach (string row in dic.Value)
                {
                    int binIdx = int.Parse(row.Split('_')[0]);
                    int eqIdx = int.Parse(row.Split('_')[1]);
                    int[] eqCntAry = new int[siteCount];

                    foreach (KeyValuePair<int, List<AdjustVddBinningRow>> adjustVddBinningRow in AdjustVddBinningRows)
                    {
                        for (int site = 0; site < siteCount; site++)
                        {
                            var rows = adjustVddBinningRow.Value.Where(x => x.PowerName.EqualsIgnoreCase(dic.Key) && x.Site == site).ToList();
                            if (rows.Exists(x => x.Type.EqualsIgnoreCase("EQN")) &&
                                rows.Exists(x => x.Type.EqualsIgnoreCase("BinCut")))
                            {
                                if (rows.Find(x => x.Type.EqualsIgnoreCase("EQN"))!.Value == eqIdx &&
                                    rows.Find(x => x.Type.EqualsIgnoreCase("BinCut"))!.Value == binIdx)
                                {
                                    if (IsBinoutDice(site, rows.FirstOrDefault()!.XCoor, rows.FirstOrDefault()!.YCoor))
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        eqCntAry[site]++;
                                    }
                                }
                            }
                        }
                    }

                    List<string> dataList =
                    [
                        dic.Key,
                        $"BinCut{binIdx} EQ{eqIdx}",
                        eqCntAry.Sum().ToString(CultureInfo.InvariantCulture)
                    ];
                    for (int site = 0; site < siteCount; site++)
                    {
                        dataList.Add(eqCntAry[site].ToString(CultureInfo.InvariantCulture));
                    }

                    data.Add(dataList);
                }
            }

            xlWorkSheet.Cells[currRow, 1].PrintExcelRowByList(data);
            xlWorkSheet.MergeColumn(1);
            xlWorkSheet.Cells.TryAutoFitColumns();
        }

        private Dictionary<string, List<string>> GetpowerNamesVsBinEq()
        {
            var outDictionary = new Dictionary<string, List<string>>();
            var powerNames = AdjustVddBinningRows.SelectMany(x => x.Value).Select(x => x.PowerName).Distinct().ToList();
            var binVsEq = new List<string>();
            int siteCount = AllDiceInfos.Count;
            foreach (KeyValuePair<int, List<AdjustVddBinningRow>> adjustVddBinningRow in AdjustVddBinningRows)
            {
                for (int pwrIdx = 0; pwrIdx < powerNames.Count; pwrIdx++)
                {
                    for (int site = 0; site < siteCount; site++)
                    {
                        var rows = adjustVddBinningRow.Value.Where(x => x.PowerName.EqualsIgnoreCase(powerNames[pwrIdx]) && x.Site == site).ToList();
                        if (rows.Any(x => x.Type.EqualsIgnoreCase("EQN")) &&
                            rows.Any(x => x.Type.EqualsIgnoreCase("BinCut")))
                        {
                            binVsEq.Add(rows.Find(x => x.Type.EqualsIgnoreCase("BinCut"))!.Value + "_" + rows.Find(x => x.Type.EqualsIgnoreCase("EQN"))!.Value);
                        }
                    }
                    if (!outDictionary.TryGetValue(powerNames[pwrIdx], out List<string>? value))
                    {
                        outDictionary.Add(powerNames[pwrIdx], [.. binVsEq.Distinct()]);
                    }
                    else
                    {
                        binVsEq.AddRange(value);
                        outDictionary[powerNames[pwrIdx]] = [.. binVsEq.Distinct()];
                    }
                }
            }

            foreach (KeyValuePair<string, List<string>> dic in outDictionary)
            {
                dic.Value.Sort();
            }
            return outDictionary;
        }

        private bool IsBinoutDice(int site, int xCoor, int yCoor)
        {
            //W/O test porgram. always increase site count
            if (_binTable.Count == 0)
            {
                return false;
            }

            foreach (Tuple<SiteInfo, List<PowerBinningLineRow>> oneDice in _powerBinningLines)
            {
                if (oneDice.Item1.Site == site && oneDice.Item1.XCoor == xCoor && oneDice.Item1.YCoor == yCoor)
                {
                    if (oneDice.Item1.SortBin == -1)
                    {
                        return true;
                    }

                    Tuple<int, string>? s = _binTable.Find(x => x.Item1 == oneDice.Item1.SortBin);
                    if (s == null)
                    {
                        return false;
                    }

                    if (s.Item2.EqualsIgnoreCase("F"))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }

        public void ExcelClose()
        {
            if (StringExtensions.IsOpened(DestExcel))
            {
                throw new Exception($"Please close file {DestExcel} !!!");
            }

            XlWorkBook.Worksheets[0].Select();
            XlPackage.SaveAs(new FileInfo(DestExcel));
            XlPackage.Dispose();
        }
    }
}
