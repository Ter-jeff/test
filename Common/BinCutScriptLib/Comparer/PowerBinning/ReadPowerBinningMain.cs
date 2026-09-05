using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

using BinCutScriptLib.Static;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib;
using TestPlanLib.BinCut;
using TestPlanLib.BinCut.PowerBinning;

namespace BinCutScriptLib.Comparer.PowerBinning
{
    internal static class ReadPowerBinningMain
    {
        private const string PowerBinning = "PwrBin";
        private const string PowerScreen = "PwrScreen";

        internal static void ReadNewPowerBinningSheet(ExcelWorkbook binCutWorkbook, ExcelWorkbook testplanWorkbook, string powerBinningPath, string tempFolder, Job job)
        {
            if (BinCutConfig.IsEnablePowerBinningHarvest)
            {
                var binningHavrSheetList = new List<ExcelWorksheet>();
                var sheetList = new List<string>();
                if (!string.IsNullOrEmpty(powerBinningPath))
                {
                    var powerBinningPackage = new ExcelPackage(new FileInfo(powerBinningPath));
                    binningHavrSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(powerBinningPackage.Workbook.Worksheets, BinCutConfig.PowerBinningSeqSheet));
                    sheetList = [.. powerBinningPackage.Workbook.Worksheets.Select(x => x.Name)];
                }
                if (binningHavrSheetList.Count == 0)
                {
                    binningHavrSheetList.AddRange(BinCutReadMainHelpers1.GetSheetFromTempFolder(BinCutReadMainHelpers1.FindPwrBinningSheetName(BinCutConfig.PowerBinningSeqSheet, tempFolder, job), tempFolder));
                }

                if (binningHavrSheetList.Count == 0 && testplanWorkbook != null)
                {
                    binningHavrSheetList.AddRange(GetBinCutExcelSheetList(testplanWorkbook.Worksheets, BinCutConfig.PowerBinningSeqSheet));
                }
                else if (binningHavrSheetList.Count == 0 && binCutWorkbook != null)
                {
                    binningHavrSheetList.AddRange(GetBinCutExcelSheetList(binCutWorkbook.Worksheets, BinCutConfig.PowerBinningSeqSheet));
                }

                if (sheetList.Count == 0)
                {
                    sheetList = [.. Directory.GetFiles(tempFolder, "*.*")];
                }

                if (binningHavrSheetList.Count != 0)
                {
                    bool caicosNewFormat = false;
                    ExcelWorksheet sheet = binningHavrSheetList.First();
                    if (BinCutData.PowerBinningSheetList.Count != 0)
                    {
                        caicosNewFormat = BinCutData.PowerBinningSheetList.First()
                            .Value.BinnedModeList.First()
                            .FactorDictionary.Keys.Any(x => x.Contains("AC_"));
                    }

                    PowerBinningHavrSheet? powerBinningHarvSheet;
                    if (caicosNewFormat)
                    {
                        var powerBinningSheetReader = new PowerBinningHarvSheetReaderCaicos(sheetList);
                        powerBinningHarvSheet = powerBinningSheetReader.ReadSheet(sheet);
                    }
                    else
                    {
                        var powerBinningSheetReader = new PowerBinningHarvSheetReader(sheetList);
                        powerBinningHarvSheet = powerBinningSheetReader.ReadSheet(sheet);

                    }
                    if (powerBinningHarvSheet != null)
                    {
                        BinCutData.PowerBinHavrSheet = powerBinningHarvSheet;
                    }
                }
            }
            else
            {
                var pwrbinSeqSheetList = new List<ExcelWorksheet>();
                if (!string.IsNullOrEmpty(powerBinningPath))
                {
                    var powerBinningPackage = new ExcelPackage(new FileInfo(powerBinningPath));
                    pwrbinSeqSheetList.AddRange(GetBinCutExcelSheetList(powerBinningPackage.Workbook.Worksheets, "Pwrbin_Seq"));
                }
                if (pwrbinSeqSheetList.Count == 0)
                {
                    pwrbinSeqSheetList.AddRange(BinCutReadMainHelpers1.GetSheetFromTempFolder("Pwrbin_Seq.txt", tempFolder));
                }

                if (pwrbinSeqSheetList.Count == 0 && testplanWorkbook != null)
                {
                    pwrbinSeqSheetList.AddRange(GetBinCutExcelSheetList(testplanWorkbook.Worksheets, "Pwrbin_Seq"));
                }
                else if (pwrbinSeqSheetList.Count == 0 && binCutWorkbook != null)
                {
                    pwrbinSeqSheetList.AddRange(GetBinCutExcelSheetList(binCutWorkbook.Worksheets, "Pwrbin_Seq"));
                }

                if (pwrbinSeqSheetList.Count != 0)
                {
                    ExcelWorksheet sheet = pwrbinSeqSheetList.First();
                    var pwrbinSeqSheetReader = new PwrbinSeqSheetReader();
                    BinCutData.PwrbinSeqSheet = pwrbinSeqSheetReader.ReadSheet(sheet);
                    BinCutData.PowerBinningSheetList = SortPowerBinningSheetList();
                }
            }
        }

        internal static Dictionary<string, PowerBinningSheet> SortPowerBinningSheetList()
        {
            var list = new Dictionary<string, PowerBinningSheet>(StringExtensions.IgnoreCase);
            foreach (string row in BinCutData.PwrbinSeqSheet!.Bin1List)
            {
                foreach (KeyValuePair<string, PowerBinningSheet> pair in BinCutData.PowerBinningSheetList)
                {
                    if (row.EqualsIgnoreCase(pair.Key))
                    {
                        if (!list.ContainsKey(pair.Key))
                        {
                            list.Add(pair.Key, pair.Value);
                        }
                    }
                }
            }
            foreach (string row in BinCutData.PwrbinSeqSheet.BinXList)
            {
                foreach (KeyValuePair<string, PowerBinningSheet> pair in BinCutData.PowerBinningSheetList)
                {
                    if (row.EqualsIgnoreCase(pair.Key))
                    {
                        if (!list.ContainsKey(pair.Key))
                        {
                            list.Add(pair.Key, pair.Value);
                        }
                    }
                }
            }
            foreach (KeyValuePair<string, PowerBinningSheet> pair in BinCutData.PowerBinningSheetList)
            {
                if (!list.ContainsKey(pair.Key))
                {
                    if (!list.ContainsKey(pair.Key))
                    {
                        list.Add(pair.Key, pair.Value);
                    }
                }
            }
            BinCutData.PowerBinningSheetList = list;
            return list;
        }

        internal static List<ExcelWorksheet> GetBinCutExcelSheetList(ExcelWorksheets excelWorksheets, string sheetName)
        {
            if (excelWorksheets.Any(x => x.Name.StartsWithIgnoreCase(sheetName)))
            {
                return [.. excelWorksheets.Where(x => x.Name.StartsWithIgnoreCase(sheetName))];
            }

            return [];
        }

        internal static void StartPowerBinning(string powerBinningPath, string tempFolder, ExcelWorkbook excelWorkbook, Action<string, Color> richTextBoxAppend)
        {
            var powerBinningSheetList = new List<ExcelWorksheet>();
            if (!string.IsNullOrEmpty(powerBinningPath))
            {
                var powerBinningPackage = new ExcelPackage(new FileInfo(powerBinningPath));
                powerBinningSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(powerBinningPackage.Workbook.Worksheets, PowerBinning));
                powerBinningSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(powerBinningPackage.Workbook.Worksheets, "Power_Binning"));
                powerBinningSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(powerBinningPackage.Workbook.Worksheets, "FinalizedSheet_toTE"));
            }
            if (powerBinningSheetList.Count == 0)
            {
                powerBinningSheetList.AddRange(BinCutReadMainHelpers1.GetSheetFromTempFolder(PowerBinning + "*.txt", tempFolder));
            }

            if (BinCutConfig.IsEnablePowerBinningHarvest)
            {
                powerBinningSheetList.AddRange(BinCutReadMainHelpers1.GetSheetFromTempFolder(PowerScreen + "*.txt", tempFolder));
            }

            if (powerBinningSheetList.Count == 0 && excelWorkbook != null)
            {
                powerBinningSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(excelWorkbook.Worksheets, PowerBinning));
                powerBinningSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(excelWorkbook.Worksheets, "Power_Binning"));
                powerBinningSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(excelWorkbook.Worksheets, "FinalizedSheet_toTE"));
            }
            if (powerBinningSheetList.Count != 0)
            {
                richTextBoxAppend("Reading Power_Binning ...", Color.Blue);
                powerBinningSheetList = [.. powerBinningSheetList
                    .GroupBy(x => x.Name)
                    .Select(x => x.First())];
                BinCutConfig.PowerBiningMappingTable = BinCutReadMainHelpers.GetPowerBiningOtherRailDic();
                foreach (ExcelWorksheet sheet in powerBinningSheetList)
                {
                    var powerBinningdReader = new PowerBinningSheetReaderCebu();
                    PowerBinningSheetCebu? powerBinningSheet = powerBinningdReader.ReadSheet(sheet);
                    if (!BinCutData.PowerBinningSheetList.ContainsKey(sheet.Name) && powerBinningSheet != null)
                    {
                        BinCutData.PowerBinningSheetList.Add(sheet.Name, powerBinningSheet);
                    }
                }
            }
        }

    }
}
