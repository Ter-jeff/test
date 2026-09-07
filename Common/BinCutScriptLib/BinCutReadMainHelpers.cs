using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.SetFunction.SetStartStep;
using BinCutScriptLib.Static;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib;
using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutConfig;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.DataStruct;
using TestPlanLib.Static;

namespace BinCutScriptLib
{
    internal static class BinCutReadMainHelpers
    {
        internal static void GetSheet(string tempFolder, string file, string sheetName, ExcelPackage excelPackage)
        {
            if (File.Exists(Path.Combine(tempFolder, file + ".txt")))
            {
                ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                using var sr = new StreamReader(Path.Combine(tempFolder, file + ".txt"));
                int rowIndex = 0;
                while (!sr.EndOfStream)
                {
                    string? line = sr.ReadLine();
                    rowIndex++;
                    if (line != null)
                    {
                        string[] lineSpt = line.Split(['\t'], StringSplitOptions.None);
                        sheet.Cells[rowIndex, 1].PrintExcelRow(lineSpt);
                    }
                }
            }
        }

        internal static TestSettingData? GetCurrentTestSetting(List<TestSettingData> testSettingDatas, EnumJob enumJob)
        {
            TestSettingData? sheet = testSettingDatas.Find(x => x.Job.EqualsIgnoreCase(enumJob.ToString()));
            if (sheet == null && enumJob.ToString().StartsWith("CP"))
            {
                sheet = testSettingDatas.Find(x => x.Job.EqualsIgnoreCase("CP1"));
            }

            if (sheet == null && enumJob.ToString().StartsWith("FT"))
            {
                sheet = testSettingDatas.Find(x => x.Job.EqualsIgnoreCase("FT1"));
            }

            sheet ??= testSettingDatas.Find(x => x.Job.EqualsIgnoreCase("CP1"));

            return sheet;
        }

        public static Dictionary<string, string> GetPowerBiningOtherRailDic()
        {
            //MSC601 vs VDD_CPU_SRAM_MS601
            var dic = new Dictionary<string, string>();
            BinningTable othRailRef = BinCutData.OtherRailTables[0];
            int index = othRailRef.ModeIdx;
            for (int rowIdx = 0; rowIdx < othRailRef.Rows.Count; rowIdx++)
            {
                string peUseOnly = othRailRef.Rows[rowIdx].RowData[index];
                if (!dic.ContainsKey(peUseOnly))
                {
                    string domain = othRailRef.Rows[rowIdx].RowData[othRailRef.DomainIdx];
                    string mode = othRailRef.Rows[rowIdx].RowData[othRailRef.ModeIdx];
                    if (BinCutConfig.DomainInOtherRail2Power.ContainsKey(domain))
                    {
                        if (BinCutConfig.DomainInOtherRail2Power[domain].EndsWithIgnoreCase(mode))
                        {
                            dic.Add(peUseOnly, BinCutConfig.DomainInOtherRail2Power[domain]);
                        }
                        else
                        {
                            dic.Add(peUseOnly, BinCutConfig.DomainInOtherRail2Power[domain] + "_" + mode);
                        }
                    }
                }
            }
            return dic;
        }

        public static List<InterpolatioNode> GetInterpolationListDic(BinningTables binningTables)
        {
            var newDic = new List<InterpolatioNode>();
            for (int i = 0; i < binningTables.Count; i++)
            {
                if (binningTables[i].IntMfIdx == -1)
                {
                    continue;
                }

                foreach (BinningRow row in binningTables[i].Rows)
                {
                    string power = "VDD_" + row.RowData[binningTables[i].DomainIdx] + "_" + row.RowData[binningTables[i].ModeIdx];

                    if (double.TryParse(row.RowData[binningTables[i].IntMfIdx], out double value))
                    {
                        if (value != 0 & !newDic.Exists(x => x.Mode.EqualsIgnoreCase(power)))
                        {
                            var it = new InterpolatioNode
                            {
                                Tableidx = i,
                                Mode = power,
                                Ratio = value,
                                Offset = binningTables[i].IntOffsetIdx != -1 && double.TryParse(row.RowData[binningTables[i].IntOffsetIdx], out value) ? value : 0,
                                Bskip = binningTables[i].IntSkipTestIdx != -1 && row.RowData[binningTables[i].IntSkipTestIdx].EqualsIgnoreCase("yes")
                            };
                            newDic.Add(it);
                        }
                    }
                }
            }
            return newDic;
        }

        internal static void GenVoltageTable(InstanceSheet instanceSheet, TestSettingData testSettingData, out VoltageTable nvTable, out VoltageTable vrsTable)
        {
            nvTable = new VoltageTable(instanceSheet);
            vrsTable = new VoltageTable(instanceSheet);
            var newTestSettingStringNv = new Dictionary<string, string>();
            var newTestSettingStringVrs = new Dictionary<string, string>();
            var categories = testSettingData.DcCategorys.Where(x => x.CategoryName.StartsWithIgnoreCase("BinCut_")).ToList();
            foreach (DcCategoryName category in categories)
            {
                var list1 = new List<Tuple<string, string>>();
                var list2 = new List<Tuple<string, string>>();
                for (int i = 0; i < BinCutData.PowerPins.Count; i++)
                {
                    string value1 = testSettingData.GetPinValue(BinCutData.PowerPins[i], category.CategoryName, category.ValueType) ?? "";
                    list1.Add(new Tuple<string, string>(BinCutData.PowerPins[i], value1));

                    string? value2 = testSettingData.GetPinValue(BinCutData.PowerPins[i] + "_Valt", category.CategoryName, category.ValueType);
                    if (string.IsNullOrEmpty(value2))
                    {
                        value2 = value1;
                    }

                    list2.Add(new Tuple<string, string>(BinCutData.PowerPins[i], value2));
                }
                string name = (category.CategoryName + "_" + category.GetvoltageType()).ToUpper();
                if (!newTestSettingStringNv.ContainsKey(name))
                {
                    newTestSettingStringNv.Add(name, BinCutData.GetNewTestSettingString(list1));
                }

                if (!newTestSettingStringVrs.ContainsKey(name))
                {
                    newTestSettingStringVrs.Add(name, BinCutData.GetNewTestSettingString(list2));
                }
            }
            nvTable.NewTestSettingString = newTestSettingStringNv;
            vrsTable.NewTestSettingString = newTestSettingStringVrs;
        }

        internal static ExcelPackage? GetBinCutExcelPackage(bool isCsharp, string filePath, string tempFolder)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                return new ExcelPackage(new FileInfo(filePath));
            }

            var excelPackage = new ExcelPackage();

            (string Config, string DefaultName, string Target)[] sheetConfigs = isCsharp ?
                [
                    (Config: BinCutConfig.BincutAteConditionNon, DefaultName: "bincut_ate_condition_non", Target: NeededSheets.BcFlow),
                    (Config: BinCutConfig.BincutEqnAppA, DefaultName: "bincut_eqn_appA", Target: NeededSheets.Binning),
                    (Config: BinCutConfig.VddBinningDefAppA2, DefaultName: "Vdd_Binning_Def_appA_2", Target: NeededSheets.BinningBinX),
                    (Config: BinCutConfig.VddBinningDefAppA3, DefaultName: "Vdd_Binning_Def_appA_3", Target: NeededSheets.BinningBinY)
                ] :
                [
                    (Config: BinCutConfig.NonBinningRail, DefaultName: "Non_Binning_Rail", Target: NeededSheets.BcFlow),
                    (Config: BinCutConfig.VddBinningDefAppA1, DefaultName: "Vdd_Binning_Def_appA_1", Target: NeededSheets.Binning),
                    (Config: BinCutConfig.VddBinningDefAppA2, DefaultName: "Vdd_Binning_Def_appA_2", Target: NeededSheets.BinningBinX),
                    (Config: BinCutConfig.VddBinningDefAppA3, DefaultName: "Vdd_Binning_Def_appA_3", Target: NeededSheets.BinningBinY)
                ];

            foreach ((string config, string defaultName, string target) in sheetConfigs)
            {
                string fileName = string.IsNullOrEmpty(config) ? defaultName : config;
                GetSheet(tempFolder, fileName, target, excelPackage);
            }

            return excelPackage.Workbook.Worksheets.Count > 0 ? excelPackage : null;
        }

        internal static List<ExcelWorksheet> GetBinCutExcelSheetList(ExcelWorksheets excelWorksheets, string sheetName)
        {
            if (excelWorksheets.Any(x => x.Name.StartsWithIgnoreCase(sheetName)))
            {
                return [.. excelWorksheets.Where(x => x.Name.StartsWithIgnoreCase(sheetName))];
            }

            return [];
        }

        internal static ExcelPackage? GetPostBinCutExcelPackage(string filePath, string tempFolder)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                return new ExcelPackage(new FileInfo(filePath));
            }

            var excelPackage = new ExcelPackage();
            string[] files = Directory.GetFiles(tempFolder, "Non_Binning_Rail_Outside_*.txt", SearchOption.AllDirectories);
            foreach (string file in files)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                if (string.IsNullOrEmpty(fileName))
                {
                    continue;
                }

                string str = fileName.Replace("Non_Binning_Rail_Outside_", "");
                string sheetName = string.IsNullOrEmpty(str) ? NeededSheets.BcFlowPost : Combination.CombineByUnderLine(NeededSheets.BcFlowPost, str);
                GetSheet(tempFolder, fileName, sheetName, excelPackage);
            }

            if (excelPackage.Workbook.Worksheets.Count > 0)
            {
                return excelPackage;
            }

            return null;
        }

        internal static void LoadBinCutConfig()
        {
            BinCutConfigXml binCutConfig = BinCutConfig.ProjectConfig.GetBinCutConfig();
            BinCutConfig.PowerType = binCutConfig.PowerTypeDic;
            BinCutConfig.IdsNames = binCutConfig.IdsNameList;
            BinCutConfig.DomainInOtherRail2Power = binCutConfig.DomainInOtherRail2PowerDic;
            BinCutConfig.PowerBinningConfig = binCutConfig.PowerBinningConfig;
            BinCutConfig.SafeVoltageFollowPayload = binCutConfig.SafeVoltageFollowPayload;
        }

        internal static void GenNvVrsTable(List<TestSettingData> testSettingDatas, Job job, bool? isVoltageByProgram, string tempFolder)
        {
            InstanceSheet instanceSheet = BinCutData.TestInstanceSheet!;
            List<Tuple<string, string>> dcSpecs = instanceSheet.GetUsedDcSpecs();
            if (isVoltageByProgram == true)
            {
                BinCutReadMainHelpers1.GetDcSpecs(instanceSheet, dcSpecs, job.JobType, tempFolder);
                return;
            }

            if (testSettingDatas == null)
            {
                BinCutReadMainHelpers1.GetDcSpecs(instanceSheet, dcSpecs, job.JobType, tempFolder);
            }
            else
            {
                TestSettingData? testSetting = GetCurrentTestSetting(testSettingDatas, job.JobType);
                if (testSetting != null &&
                    testSetting.DcCategorys.Exists(x => x.CategoryName.StartsWithIgnoreCase("BinCut_")))
                {
                    GenVoltageTable(instanceSheet, testSetting, out VoltageTable? nvTable, out VoltageTable? vrsTable);
                    BinCutData.NvTable = nvTable;
                    BinCutData.VrsTable = vrsTable;
                }
                else
                {
                    BinCutReadMainHelpers1.GetDcSpecs(instanceSheet, dcSpecs, job.JobType, tempFolder);
                }
            }
        }

        internal static void ReadPatternDashboard(string patternDashBoardFilePath)
        {
            // Reading pattern dashboard
            if (!string.IsNullOrEmpty(patternDashBoardFilePath))
            {
                Dictionary<string, PatternData>? patternDic = PatternListReader.GetPatternListDic(patternDashBoardFilePath);
                BinCutData.PatList = patternDic != null ? [.. patternDic.Select(x => x.Value)] : [];
                if (patternDic != null)
                {
                    BinCutData.HasPatList = true;
                }
            }
        }

        internal static BinCutFlowTables GetBinCutFlowTables(ExcelWorkbook excelWorkbook, bool isCsharp, List<string> binningTitleList, Dictionary<string, List<string>> domainDic, Job job, string tempFolder)
        {
            string selectedConfigValue = isCsharp ? BinCutConfig.BincutAteConditionNon : BinCutConfig.NonBinningRail;
            string defaultValue = isCsharp ? "bincut_ate_condition_non" : "Non_Binning_Rail";
            ExcelWorksheet binCutWorksheet = excelWorkbook != null ?
                excelWorkbook.Worksheets[NeededSheets.BcFlow] ?? excelWorkbook.Worksheets[NeededSheets.BcAte] : BinCutReadMainHelpers1.GetSheetFromTempFolder(selectedConfigValue, tempFolder).First();
            var binCutSheetReader = new BinCutFlowSheetReader(binningTitleList, domainDic);
            #region overwrite bincut_ate_condition_non (nonbinning rail) with jobName if nonbinning rail use default value
            if (selectedConfigValue == defaultValue &&
                File.Exists(Path.Combine(tempFolder, selectedConfigValue + "_" + job.JobName + ".txt")))
            {
                binCutWorksheet = BinCutReadMainHelpers1.GetSheetFromTempFolder(selectedConfigValue + "_" + job.JobName + ".txt", tempFolder).First();
            }
            #endregion
            BinCutFlowTables? binCutFlowTables = binCutSheetReader.ReadSheet(binCutWorksheet);
            if (binCutFlowTables != null)
            {
                foreach (BinCutFlowTable binCutFlowTable in binCutFlowTables)
                {
                    binCutFlowTable.AddToErrorReport();
                }
            }

            BinCutFlowTables result = binCutFlowTables ?? [];
            BinCutData.BinCutFlowTables = result;

            return result;
        }

        internal static void LoadBinCutPostFlowSheet(bool isCsharp, List<string> binningTitleList, Dictionary<string, List<string>> domainDic, string binCutPostFilePath, string tempFolder)
        {
            ExcelPackage? postBinCutExcelPackage = GetPostBinCutExcelPackage(binCutPostFilePath, tempFolder);
            ExcelWorkbook? binCutPostWorkbook = postBinCutExcelPackage?.Workbook;

            var binCutPostWorksheet = new List<ExcelWorksheet>();
            if (binCutPostWorkbook != null)
            {
                IEnumerable<ExcelWorksheet> postSheets = binCutPostWorkbook.Worksheets.Where(sheet =>
                    sheet.Name.StartsWithIgnoreCase(NeededSheets.BcFlowPost)
                );
                binCutPostWorksheet.AddRange(postSheets);
            }
            else
            {
                List<string> outsideBinCutList = [.. Directory.GetFiles(tempFolder, (isCsharp ? "bincut_ate_condition_outside_*" : "Non_Binning_Rail_Outside_*") + ".txt")];
                foreach (string fileName in outsideBinCutList)
                {
                    string file = fileName.Replace(tempFolder + Path.DirectorySeparatorChar, "");
                    if (BinCutReadMainHelpers1.GetSheetFromTempFolder(file, tempFolder).Count > 0)
                    {
                        binCutPostWorksheet.Add(BinCutReadMainHelpers1.GetSheetFromTempFolder(file, tempFolder).First());
                    }
                }
            }
            if (binCutPostWorksheet.Count == 0)
            {
                return;
            }
            foreach (ExcelWorksheet sheet in binCutPostWorksheet)
            {
                var binCutPostSheetReader = new BinCutFlowSheetReader(binningTitleList, domainDic);
                BinCutFlowTables? postFlowTables = binCutPostSheetReader.ReadSheet(sheet);
                if (postFlowTables == null)
                {
                    continue;
                }
                BinCutData.PostFlowSheets.Add(postFlowTables);
            }
        }
    }
}
