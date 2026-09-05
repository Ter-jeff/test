using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.BistBira.IgxlSheet;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;

using OfficeOpenXml;

using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.BinCut;
using TestPlanLib.DataStruct;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.Basic.Business.GenMappingTable
{
    internal class NewDsscMappingTable
    {
        public SelsrmMappingSheet SelsrmMappingSheet = new SelsrmMappingSheet();
        public List<BistCzDsscMappingTable> BistCzDsscMappingTables = new List<BistCzDsscMappingTable>();
        public static Dictionary<string, PatternData> PatSetTimeSetDictionary = new Dictionary<string, PatternData>();
        private Dictionary<string, Dictionary<string, BistCzDsscCategory>> _bistCzDsscCategoryDic = new Dictionary<string, Dictionary<string, BistCzDsscCategory>>();

        public void Workflow(ExcelWorksheet sheet, string outputFolder, MultiTestSettingSheetsSingleton testSetting)
        {
            try
            {
                var selsrmMappingSheetReader = new SelsrmMappingSheetReader();
                SelsrmMappingSheet = selsrmMappingSheetReader.ReadSheet(sheet);
                PatSetTimeSetDictionary = AcTSetCategoryMapSingleton.Instance().PatternList;

                Precheck(SelsrmMappingSheet, testSetting);

                var neededSheetsFromConfig = new List<string>
                {
                    NeededSheets.ScanScghCpu,
                    NeededSheets.ScanScghGpu,
                    NeededSheets.ScanScghSoc,
                    NeededSheets.MbistCharScgCpu,
                    NeededSheets.MbistCharScgGpu,
                    NeededSheets.MbistCharScgSoc
                };
                var scghSheets = new List<string>();
                foreach (string neededSheet in neededSheetsFromConfig)
                {
                    if (neededSheet.Contains(","))
                    {
                        scghSheets.AddRange(neededSheet.Split(',').ToList());
                    }
                    else
                    {
                        scghSheets.Add(neededSheet);
                    }
                }
                var groups = testSetting.TestSettingSheetsList.GroupBy(x => x.OriginJob).ToList();
                foreach (IGrouping<string, TestSettingData> group in groups)
                {
                    var bistCzDsscOnePatterns = new List<BistCzDsscOnePattern>();
                    foreach (string scghSheet in scghSheets)
                    {
                        var prodCharSheet = new ProdCharSheet();
                        ExcelWorksheet scgWorksheet = EpWorkbook.ScghWorkbook.Worksheets[scghSheet];
                        if (scgWorksheet != null)
                        {
                            var sheetReader = new ProdCharSheetReader();
                            prodCharSheet.RowList.AddRange(sheetReader.ReadSheet(scgWorksheet).RowList);
                        }
                        TestSettingData testSettingSheet = group.First();

                        string currentJob = testSettingSheet.SheetName.Split('_').Last();
                        _bistCzDsscCategoryDic = GetBistCzDsscCategoryDic(testSettingSheet);
                        bistCzDsscOnePatterns.AddRange(GetCzDssc(prodCharSheet, scghSheet, testSettingSheet, SelsrmMappingSheet, currentJob));
                    }
                    foreach (TestSettingData row in group)
                    {
                        string czDsscName = "DSSCMappingTable" + "_" + row.Job;
                        var bistCzDsscMappingTable = new BistCzDsscMappingTable();
                        string outputFile = Path.Combine(outputFolder, czDsscName + ".txt");
                        bistCzDsscMappingTable.OutputFile = outputFile;
                        bistCzDsscMappingTable.BistCzDsscOnePatterns.AddRange(bistCzDsscOnePatterns);
                        BistCzDsscMappingTables.Add(bistCzDsscMappingTable);
                    }
                }

                foreach (BistCzDsscMappingTable table in BistCzDsscMappingTables)
                {
                    if (table.BistCzDsscOnePatterns.Count == 0)
                    {
                        continue;
                    }

                    WriteTxt(GenDictionary(table), table.OutputFile);

                    TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, Path.GetFileNameWithoutExtension(table.OutputFile));
                }
            }
            catch (Exception e)
            {
                Response.Report("Meet an Error in CzDsscMapping: " + e.StackTrace, EnumMessageLevel.Error, 100);
            }
        }

        public void Precheck(SelsrmMappingSheet selsrmMappingSheet, MultiTestSettingSheetsSingleton testSetting)
        {
            var newgroups = testSetting.TestSettingSheetsList.GroupBy(x => x.OriginJob).ToList();
            foreach (IGrouping<string, TestSettingData> tmp in newgroups)
            {
                TestSettingData testSettingSheetNew = tmp.First();
                List<SelsrmMappingTableRow> selsrmMappingTableRows = GetSelsrmMappingTableRows(selsrmMappingSheet, testSettingSheetNew.Job);
                var dictionary = selsrmMappingTableRows.GroupBy(x => x.Pattern).ToDictionary(x => x.Key, x => x.ToList());
                foreach (KeyValuePair<string, List<SelsrmMappingTableRow>> pair in dictionary)
                {
                    foreach (SelsrmMappingTableRow row in pair.Value)
                    {
                        var logicPins = row.LogicPins.Split(',').Select(x => x.Trim()).Where(y => !string.IsNullOrEmpty(y)).ToList();
                        var sramPins = row.SramPins.Split(',').Select(x => x.Trim()).Where(y => !string.IsNullOrEmpty(y)).ToList();
                        bool allPass = true;
                        foreach (string logicPin in logicPins)
                        {
                            List<TestSettingRow> pinValtInCategory = testSettingSheetNew.GetAllCategory(new List<string> { logicPin + "_Valt" });
                            if (!pinValtInCategory.Any())
                            {
                                string mErrorMessage = $"Logic Pin: {logicPin} do not have the *_Valt voltage in {testSettingSheetNew.Job}";
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_03, selsrmMappingSheet.SheetName, row.RowNum, 0, $"Logic Pin: {logicPin} do not have the *_Valt voltage in {testSettingSheetNew.Job}", new string[] { logicPin, testSettingSheetNew.Job });
                                allPass = false;
                            }
                        }
                        foreach (string sramPin in sramPins)
                        {
                            List<TestSettingRow> pinValtInCategory = testSettingSheetNew.GetAllCategory(new List<string> { sramPin + "_Valt" });
                            if (!pinValtInCategory.Any())
                            {
                                string mErrorMessage = $"Sram Pin: {sramPin} do not have the *_Valt voltage in {testSettingSheetNew.Job}";
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_04, selsrmMappingSheet.SheetName, row.RowNum, 0, $"Sram Pin: {sramPin} do not have the *_Valt voltage in {testSettingSheetNew.Job}", new string[] { sramPin, testSettingSheetNew.Job });
                                allPass = false;
                            }
                        }

                        if (allPass)
                        {
                            if (row.Selsrm0 == row.Selsrm1)
                            {
                                //error1
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_05, selsrmMappingSheet.SheetName, row.RowNum, 0, row.LogicPins + " and " + row.SramPins + " should have the different source bit", new string[] { row.LogicPins, row.SramPins });
                            }
                        }
                        if (row.LogicPins.IndexOf("PRESERVED", StringComparison.OrdinalIgnoreCase) != -1 || row.SramPins.IndexOf("PRESERVED", StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            if (row.Selsrm0 != row.Selsrm1)
                            {
                                //error3
                                ErrorReportManager.AddError(BasicErrorType.E_FormatError_06, selsrmMappingSheet.SheetName, row.RowNum, 0, "The preserved pin should have the same source bit");
                            }
                        }
                    }
                }
            }
        }

        private Dictionary<string, List<BistCzDsscCategory>> GenDictionary(BistCzDsscMappingTable bistCzDsscOnePatterns)
        {
            var dictionary = new Dictionary<string, List<BistCzDsscCategory>>();
            foreach (BistCzDsscOnePattern bistCzDsscOnePattern in bistCzDsscOnePatterns.BistCzDsscOnePatterns)
            {
                if (
                    !dictionary.ContainsKey(bistCzDsscOnePattern.BistCzDsscCategories.First().BistCzDsscCategoryRows.First().PatternName))
                {
                    dictionary.Add(bistCzDsscOnePattern.BistCzDsscCategories.First().BistCzDsscCategoryRows.First().PatternName, bistCzDsscOnePattern.BistCzDsscCategories);
                }
                else
                {
                    dictionary[bistCzDsscOnePattern.BistCzDsscCategories.First().BistCzDsscCategoryRows.First().PatternName] = dictionary[bistCzDsscOnePattern.BistCzDsscCategories.First().BistCzDsscCategoryRows.First().PatternName]
                            .Concat(bistCzDsscOnePattern.BistCzDsscCategories).ToList();
                }
            }

            return dictionary;
        }

        public void WriteHeader(StreamWriter sw, List<BistCzDsscCategory> tmpCategories)
        {
            var header = new StringBuilder();
            if (tmpCategories.Count > 0)
            {
                header.Append("\t" + "Pattern Name" + "\t" + "Cycle Number" + "\t" + "Segment Number" + "\t" + "Description" + "\t" + "Test_AutoSwitch");
                {
                    foreach (BistCzDsscCategory item in tmpCategories)
                    {
                        header.Append("\t" + item.CategoryName);
                    }
                }
            }
            sw.WriteLine(header.ToString().StartsWith("\t")
                ? header.ToString().Substring(1)
                : header.ToString());
        }

        public void WriteContent(StreamWriter sw, List<string> pattern, List<BistCzDsscCategory> tmpCategories)
        {
            if (pattern.Count == 0)
            {
                pattern.Add(tmpCategories.First().BistCzDsscCategoryRows.First().PatternName);
            }

            int max = tmpCategories.Max(x => x.BistCzDsscCategoryRows.Count);
            if (max > pattern.Count)
            {
                for (int i = 0; i < max; i++)
                {
                    var data = new StringBuilder();

                    for (int index = 0; index < tmpCategories.Count; index++)
                    {
                        BistCzDsscCategory category = tmpCategories[index];

                        if (i >= category.BistCzDsscCategoryRows.Count && index == 0)
                        {
                            data.Append("\t" + "\t" + "\t" + "\t" + "\t");
                        }
                        else if (i < pattern.Count && index == 0)
                        {
                            BistCzDsscCategoryRow row = tmpCategories.First().BistCzDsscCategoryRows[i];
                            data.Append("\t" + pattern[i] + "\t" + row.CycleNumber + "\t" +
                                        row.SegmentNumber + "\t" + row.Description + "\t" +
                                        row.TestAutoSwitch);
                        }
                        else if (index == 0)
                        {
                            BistCzDsscCategoryRow row = tmpCategories.First().BistCzDsscCategoryRows[i];
                            data.Append("\t" + "\t" + row.CycleNumber + "\t" + row.SegmentNumber + "\t" +
                                        row.Description + "\t" + row.TestAutoSwitch);
                        }

                        if (i >= category.BistCzDsscCategoryRows.Count)
                        {
                            data.Append("\t");
                        }
                        else
                        {
                            if (category.BistCzDsscCategoryRows[i].Values == "-1")
                            {
                                category.BistCzDsscCategoryRows[i].Values = "";
                            }

                            data.Append("\t" + category.BistCzDsscCategoryRows[i].Values);
                        }
                    }

                    sw.WriteLine(data.ToString().StartsWith("\t") ? data.ToString().Substring(1) : data.ToString());

                    if (i == max - 1)
                    {
                        sw.WriteLine("");
                    }
                }
            }
            else
            {
                for (int i = 0; i < pattern.Count; i++)
                {
                    var data = new StringBuilder();

                    for (int index = 0; index < tmpCategories.Count; index++)
                    {
                        BistCzDsscCategory category = tmpCategories[index];

                        if (i >= category.BistCzDsscCategoryRows.Count && index == 0)
                        {
                            data.Append("\t" + pattern[i] + "\t" + "\t" + "\t" + "\t");
                        }
                        else if (i < max && index == 0)
                        {
                            BistCzDsscCategoryRow row = tmpCategories.First().BistCzDsscCategoryRows[i];
                            data.Append("\t" + pattern[i] + "\t" + row.CycleNumber + "\t" + row.SegmentNumber + "\t" + row.Description + "\t" + row.TestAutoSwitch);
                        }
                        else if (index == 0)
                        {
                            data.Append("\t" + pattern[i] + "\t" + "\t" + "\t" + "\t" + "\t" + "\t" + "\t");
                        }

                        if (i >= category.BistCzDsscCategoryRows.Count)
                        {
                            data.Append("\t");
                        }
                        else
                        {
                            if (category.BistCzDsscCategoryRows[i].Values == "-1")
                            {
                                category.BistCzDsscCategoryRows[i].Values = "";
                            }

                            data.Append("\t" + category.BistCzDsscCategoryRows[i].Values);
                        }
                    }

                    sw.WriteLine(data.ToString().StartsWith("\t")
                        ? data.ToString().Substring(1)
                        : data.ToString());

                    if (i == pattern.Count - 1)
                    {
                        sw.WriteLine("");
                    }
                }
            }

            pattern.Clear();
            tmpCategories.Clear();
        }

        public void WriteTxt(Dictionary<string, List<BistCzDsscCategory>> bistCzDsscCategoryDictionary, string outputFile)
        {
            var dictionary = bistCzDsscCategoryDictionary.OrderBy(x => x.Value.Count).ThenBy(x => x.Value.First().BistCzDsscCategoryRows.First().Values).ToDictionary(x => x.Key, x => x.Value);

            using (var sw = new StreamWriter(outputFile))
            {
                var bistCzDsscCategories = new List<BistCzDsscCategory>();
                var pattern = new List<string>();
                foreach (KeyValuePair<string, List<BistCzDsscCategory>> categories in dictionary)
                {
                    if (bistCzDsscCategories.Count == 0)
                    {
                        bistCzDsscCategories.AddRange(categories.Value);
                    }

                    if (
                        bistCzDsscCategories.SelectMany(x => x.CategoryName)
                            .SequenceEqual(categories.Value.SelectMany(x => x.CategoryName))
                        && bistCzDsscCategories.First().BistCzDsscCategoryRows.SelectMany(x => x.Values)
                            .SequenceEqual(categories.Value.First().BistCzDsscCategoryRows.SelectMany(x => x.Values))
                        && bistCzDsscCategories.First().BistCzDsscCategoryRows.First().Description
                            .Equals(categories.Value.First().BistCzDsscCategoryRows.First().Description))
                    {
                        if (pattern.Count == 0 &&
                            !bistCzDsscCategories.First()
                                .BistCzDsscCategoryRows.First()
                                .PatternName.Equals(categories.Key, StringComparison.OrdinalIgnoreCase))
                        {
                            pattern.Add(bistCzDsscCategories.First().BistCzDsscCategoryRows.First().PatternName);
                        }
                        pattern.Add(categories.Key);
                    }
                    else
                    {
                        WriteHeader(sw, bistCzDsscCategories);
                        WriteContent(sw, pattern, bistCzDsscCategories);
                        bistCzDsscCategories.AddRange(categories.Value);
                    }
                }

                WriteHeader(sw, bistCzDsscCategories);
                WriteContent(sw, pattern, bistCzDsscCategories);
            }
        }

        private Dictionary<string, Dictionary<string, BistCzDsscCategory>> GetBistCzDsscCategoryDic(TestSettingData testSettingSheet)
        {

            var bistCzDsscCategoryRowsDic = new Dictionary<string, Dictionary<string, BistCzDsscCategory>>();
            List<DcCategoryName> dcCategoryNames = testSettingSheet.DcCategorys;
            if (dcCategoryNames != null && dcCategoryNames.Count != 0)
            {
                List<SelsrmMappingTableRow> selsrmMappingTableRows = GetSelsrmMappingTableRows(SelsrmMappingSheet, testSettingSheet.OriginJob);
                var dictionary = selsrmMappingTableRows.GroupBy(x => x.Pattern).ToDictionary(x => x.Key, x => x.ToList());
                foreach (KeyValuePair<string, List<SelsrmMappingTableRow>> pair in dictionary)
                {
                    var bistCzDsscCategoryDic = new Dictionary<string, BistCzDsscCategory>();
                    var bistCzDsscCategoryRows = new List<BistCzDsscCategoryRow>();
                    var extendValue = new List<SelsrmMappingTableRow>();
                    if (pair.Value.First().Bits.Equals("NA", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (pair.Value.Count < Convert.ToInt32(pair.Value.Max(x => x.Bits)))
                    {
                        int i = 0;
                        foreach (SelsrmMappingTableRow row in pair.Value)
                        {
                            if (Convert.ToInt32(row.Bits) >= i) // !=?
                            {
                                for (int j = i; j <= Convert.ToInt32(row.Bits); j++)
                                {
                                    extendValue.Add(row);
                                    i++;
                                }
                            }
                        }
                    }
                    else
                    {
                        extendValue.AddRange(pair.Value);
                    }
                    foreach (SelsrmMappingTableRow row in extendValue)
                    {
                        var logicPins = new List<string>();
                        var valueLogicPinsValtAllCategory = new List<TestSettingRow>();
                        var sramPins = new List<string>();
                        var valueSramPinsValtAllCategory = new List<TestSettingRow>();
                        foreach (string pin in row.LogicPins.Split(','))
                        {
                            List<TestSettingRow> category = testSettingSheet.GetAllCategory(new List<string> { pin + "_Valt" });
                            if (!category.Any())
                            {
                                IEnumerable<string> allPins = TestProgram.IgxlWorkBk.PinMapPair.Value.GetPinsFromGroup(pin).Select(x => x.PinName).ToList();
                                category = testSettingSheet.GetAllCategory(allPins.Select(x => x.Trim() + "_Valt").ToList());
                                logicPins.AddRange(allPins);
                            }
                            else
                            {
                                logicPins.Add(pin);
                            }

                            if (category.Any())
                            {
                                valueLogicPinsValtAllCategory.AddRange(category);
                            }
                        }
                        foreach (string pin in row.SramPins.Split(','))
                        {
                            List<TestSettingRow> category = testSettingSheet.GetAllCategory(new List<string> { pin + "_Valt" });
                            if (!category.Any())
                            {
                                IEnumerable<string> allPins = TestProgram.IgxlWorkBk.PinMapPair.Value.GetPinsFromGroup(pin).Select(x => x.PinName).ToList();
                                category = testSettingSheet.GetAllCategory(allPins.Select(x => x.Trim() + "_Valt").ToList());
                                sramPins.AddRange(allPins);
                            }
                            else
                            {
                                sramPins.Add(pin);
                            }

                            if (category.Any())
                            {
                                valueSramPinsValtAllCategory.AddRange(category);
                            }
                        }
                        if (valueLogicPinsValtAllCategory.Count == logicPins.Count && valueSramPinsValtAllCategory.Count == sramPins.Count)  //any one is missing??
                        {
                            foreach (DcCategoryName dcCategoryName in dcCategoryNames)
                            {
                                var bistCzDsscCategoryRow = new BistCzDsscCategoryRow
                                {
                                    TestAutoSwitch = "S",
                                    Values = CompareLogicSramList(row, dcCategoryName.CategoryName,
                                        dcCategoryName.ValueType, valueLogicPinsValtAllCategory,
                                        valueSramPinsValtAllCategory),
                                    Description = row.Block + ":" + row.LogicPins + "_" +
                                                  row.SramPins,
                                    CategoryName = "Test_" + dcCategoryName.CategoryName + "_" +
                                                   dcCategoryName.ValueType
                                };

                                bistCzDsscCategoryRows.Add(bistCzDsscCategoryRow);
                            }
                        }
                        else
                        {
                            foreach (DcCategoryName dcCategoryName in dcCategoryNames)
                            {
                                var bistCzDsscCategoryRow = new BistCzDsscCategoryRow();

                                if (row.Selsrm0 == row.Selsrm1)
                                {
                                    bistCzDsscCategoryRow.TestAutoSwitch = row.Selsrm0;
                                    bistCzDsscCategoryRow.Values = row.Selsrm0;
                                }
                                else
                                {
                                    bistCzDsscCategoryRow.TestAutoSwitch = "S";
                                    bistCzDsscCategoryRow.Values = "Error";
                                }

                                bistCzDsscCategoryRow.Description = row.Block + ":" + row.LogicPins + "_" +
                                                                    row.SramPins;
                                bistCzDsscCategoryRow.CategoryName = "Test_" + dcCategoryName.CategoryName + "_" +
                                                                     dcCategoryName.ValueType;
                                bistCzDsscCategoryRows.Add(bistCzDsscCategoryRow);
                            }
                        }
                    }

                    if (bistCzDsscCategoryRows.Count == 0)
                    {
                        bistCzDsscCategoryRowsDic = new Dictionary<string, Dictionary<string, BistCzDsscCategory>>();
                        break;
                    }

                    foreach (IGrouping<string, BistCzDsscCategoryRow> group in bistCzDsscCategoryRows.GroupBy(x => x.CategoryName))
                    {
                        var bistCzDsscCategory = new BistCzDsscCategory
                        {
                            CategoryName = group.First().CategoryName,
                            BistCzDsscCategoryRows = group.ToList()
                        };
                        bistCzDsscCategoryDic.Add(bistCzDsscCategory.CategoryName, bistCzDsscCategory);
                    }
                    bistCzDsscCategoryRowsDic.Add(pair.Key.TrimStart('*').TrimEnd('*').Replace("*", "\\w+"), bistCzDsscCategoryDic);
                }
            }

            return bistCzDsscCategoryRowsDic;
        }

        private List<BistCzDsscOnePattern> GetCzDssc(ProdCharSheet prodCharSheet, string scghSheet, TestSettingData testSettingSheet, SelsrmMappingSheet selsrmMappingSheet, string job)
        {
            var bistCzDsscOnePatterns = new List<BistCzDsscOnePattern>();
            string domain = scghSheet.Split('_')[0];
            string block = scghSheet.Split('_')[1];
            try
            {
                List<DcCategoryName> dcCategoryNames = GetDcCategoryNames(testSettingSheet, block, domain);
                if (dcCategoryNames == null)
                {
                    string blockStr = scghSheet.Split('_')[2];
                    dcCategoryNames = GetDcCategoryNames(testSettingSheet, blockStr, domain);
                }
                if (dcCategoryNames != null && dcCategoryNames.Count != 0)
                {
                    List<SelsrmMappingTableRow> selsrmMappingTableRows = GetSelsrmMappingTableRows(selsrmMappingSheet, job);
                    List<string> dsscPatternList = GetPatternList(prodCharSheet.RowList, selsrmMappingTableRows);

                    foreach (string pattern in dsscPatternList)
                    {
                        var onePattern = new BistCzDsscOnePattern();
                        var bistCzDsscCategories = new Dictionary<string, BistCzDsscCategory>();
                        foreach (string key in _bistCzDsscCategoryDic.Keys)
                        {
                            if (Regex.IsMatch(pattern, key, RegexOptions.IgnoreCase))
                            {
                                bistCzDsscCategories = _bistCzDsscCategoryDic[key];
                                break;
                            }
                        }
                        foreach (DcCategoryName dcCategoryName in dcCategoryNames)
                        {
                            if (selsrmMappingTableRows == null || selsrmMappingTableRows.Count == 0)
                            {
                                continue;
                            }

                            string categoryName = "Test_" + dcCategoryName.CategoryName + "_" + dcCategoryName.ValueType;
                            BistCzDsscCategory bistCzDsscCategory = bistCzDsscCategories[categoryName].Copy();
                            foreach (BistCzDsscCategoryRow row in bistCzDsscCategory.BistCzDsscCategoryRows)
                            {
                                row.PatternName = pattern;
                            }

                            onePattern.BistCzDsscCategories.Add(bistCzDsscCategory);
                        }
                        bistCzDsscOnePatterns.Add(onePattern);
                    }
                }
            }
            catch (Exception e)
            {
                Response.Report("Meet an Error in GetCzDssc: " + e.StackTrace, EnumMessageLevel.Error, 100);
            }
            return bistCzDsscOnePatterns;
        }

        private static List<string> GetPatternList(List<ProdCharSheetRow> patternData, List<SelsrmMappingTableRow> selsrmMappingTableRows)
        {
            var patternName = new List<string>();

            patternName.AddRange(selsrmMappingTableRows.Select(x => x.Pattern.TrimStart('*').TrimEnd('*').Replace("*", "\\w+")).Distinct().ToList());

            var dsscPatternList = new List<string>();

            foreach (string name in patternName)
            {
                dsscPatternList.AddRange(
                    patternData.Select(x => x.PayloadValue)
                        .Where(
                            x =>
                                Regex.IsMatch(x, name, RegexOptions.IgnoreCase) &&
                                PatSetTimeSetDictionary.ContainsKey(x.ToLower()) &&
                                PatSetTimeSetDictionary[x.ToLower()].Use == "use").Distinct().ToList());

                foreach (List<string> initList in patternData.Select(x => x.InitAliasList))
                {

                    dsscPatternList.AddRange(
                        initList.FindAll(
                            x =>
                                Regex.IsMatch(x, name, RegexOptions.IgnoreCase) &&
                                PatSetTimeSetDictionary.ContainsKey(x.ToLower()) &&
                                PatSetTimeSetDictionary[x.ToLower()].Use == "use").Distinct().ToList());
                }
            }
            dsscPatternList = dsscPatternList.Distinct().ToList();
            return dsscPatternList;
        }

        internal static List<DcCategoryName> GetDcCategoryNames(TestSettingData testSettingSheet, string block, string domain)
        {

            if (block.Equals("BI", StringComparison.CurrentCultureIgnoreCase))
            {
                if (domain.Equals("C", StringComparison.OrdinalIgnoreCase))
                {
                    return testSettingSheet.DcCategorys.Where(x => x.CategoryName.StartsWith("Mbist_Cpu", StringComparison.CurrentCultureIgnoreCase)).ToList();
                }

                if (domain.Equals("L", StringComparison.OrdinalIgnoreCase))
                {
                    return testSettingSheet.DcCategorys.Where(x => x.CategoryName.StartsWith("Mbist_Gfx", StringComparison.CurrentCultureIgnoreCase) ||
                                                                   x.CategoryName.StartsWith("Mbist_Gpu", StringComparison.CurrentCultureIgnoreCase)).ToList();
                }

                if (domain.Equals("S", StringComparison.OrdinalIgnoreCase))
                {
                    return testSettingSheet.DcCategorys.Where(x => x.CategoryName.StartsWith("Mbist_Soc", StringComparison.CurrentCultureIgnoreCase)).ToList();
                }
            }
            if (block.Equals("SC", StringComparison.CurrentCultureIgnoreCase))
            {
                if (domain.Equals("C", StringComparison.OrdinalIgnoreCase))
                {
                    return testSettingSheet.DcCategorys.Where(x => x.CategoryName.StartsWith("Sa_Cpu", StringComparison.CurrentCultureIgnoreCase) ||
                                                                   x.CategoryName.StartsWith("SaChain_Cpu", StringComparison.CurrentCultureIgnoreCase) ||
                                                                   x.CategoryName.StartsWith("Td_Cpu", StringComparison.CurrentCultureIgnoreCase)).ToList();
                }

                if (domain.Equals("L", StringComparison.OrdinalIgnoreCase))

                {
                    return testSettingSheet.DcCategorys.Where(
                            x => x.CategoryName.StartsWith("Sa_Gfx", StringComparison.CurrentCultureIgnoreCase) ||
                                 x.CategoryName.StartsWith("SaChain_Gfx", StringComparison.CurrentCultureIgnoreCase) ||
                                 x.CategoryName.StartsWith("Td_Gfx", StringComparison.CurrentCultureIgnoreCase) ||
                                 x.CategoryName.StartsWith("Sa_Gpu", StringComparison.CurrentCultureIgnoreCase) ||
                                 x.CategoryName.StartsWith("SaChain_Gpu", StringComparison.CurrentCultureIgnoreCase) ||
                                 x.CategoryName.StartsWith("Td_Gpu", StringComparison.CurrentCultureIgnoreCase))
                        .ToList();
                }

                if (domain.Equals("S", StringComparison.OrdinalIgnoreCase))
                {
                    return testSettingSheet.DcCategorys.Where(x => x.CategoryName.StartsWith("Sa_Soc", StringComparison.CurrentCultureIgnoreCase) ||
                                                                   x.CategoryName.StartsWith("SaChain_Soc", StringComparison.CurrentCultureIgnoreCase) ||
                                                                   x.CategoryName.StartsWith("Td_Soc", StringComparison.CurrentCultureIgnoreCase)).ToList();
                }
            }

            return null;
        }

        internal static List<SelsrmMappingTableRow> GetSelsrmMappingTableRows(SelsrmMappingSheet selsrmMappingSheet, string job = "")
        {
            var selsrmMappingTableRows = new List<SelsrmMappingTableRow>();
            List<SelsrmMappingTableRow> rows = selsrmMappingSheet.Rows;
            foreach (SelsrmMappingTableRow row in rows)
            {
                if (!row.Block.Equals("RTOS") && !row.Block.Equals("*"))
                {
                    continue;
                }

                List<string> arr = row.Stage.Split(',').ToList();
                foreach (string searchJob in arr)
                {
                    string searchText = "^" + searchJob.Replace("*", ".*").Replace("/", "|") + "$";
                    if (Regex.IsMatch(job, searchText))
                    {
                        if (!string.IsNullOrEmpty(row.LogicPins) && !string.IsNullOrEmpty(row.SramPins))
                        {
                            selsrmMappingTableRows.Add(row);
                        }
                    }
                }
            }
            return selsrmMappingTableRows;
        }

        public string CompareLogicSramList(SelsrmMappingTableRow row, string categoryName, CategoryValueType categoryValueType, List<TestSettingRow> valueLogicPinsAllCategoryList, List<TestSettingRow> valueSramPinsAllCategoryList)
        {
            if (valueLogicPinsAllCategoryList.Count != 0 && valueSramPinsAllCategoryList.Count != 0)
            {
                List<double> logicValues = new List<double>();
                List<double> sramValues = new List<double>();
                foreach (TestSettingRow valueLogicPinsAllCategory in valueLogicPinsAllCategoryList)
                {
                    TestSettingData.GetValuebyTestSettingRow(categoryName, categoryValueType, valueLogicPinsAllCategory, out string valueLogicPin);
                    if (double.TryParse(valueLogicPin, out double logicValue))
                    {
                        logicValues.Add(logicValue);
                    }
                }
                foreach (TestSettingRow valueSramPinsAllCategory in valueSramPinsAllCategoryList)
                {
                    TestSettingData.GetValuebyTestSettingRow(categoryName, categoryValueType, valueSramPinsAllCategory, out string valueSramPin);
                    if (double.TryParse(valueSramPin, out double sramValue))
                    {
                        sramValues.Add(sramValue);
                    }
                }

                List<bool> compareResults = new List<bool>();
                foreach (double logicValue in logicValues)
                {
                    foreach (double sramValue in sramValues)
                    {
                        if (logicValue > sramValue)
                        {
                            compareResults.Add(true);
                        }
                        else
                        {
                            compareResults.Add(false);
                        }
                    }
                }

                if (logicValues.Count == valueLogicPinsAllCategoryList.Count && sramValues.Count == valueSramPinsAllCategoryList.Count)
                {
                    if (compareResults.All(x => x))
                    {
                        return row.Selsrm0;
                    }

                    if (compareResults.All(x => !x))
                    {
                        return row.Selsrm1;
                    }
                }
            }
            else if (row.LogicPins.Equals(row.SramPins, StringComparison.OrdinalIgnoreCase) && row.Selsrm0 != row.Selsrm1)
            {
                return "-1";
            }
            return "Error";
        }
    }
}
