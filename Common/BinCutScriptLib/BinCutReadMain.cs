using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Comparer;
using BinCutScriptLib.Comparer.PowerBinning;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using LogLib.Utility;

using OfficeOpenXml;

using TestPlanLib;
using TestPlanLib.Basic;
using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutConfig;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.DataStruct;

namespace BinCutScriptLib
{
    public partial class BinCutReadMain(BinCutScriptMain binCutScriptMain, Action<string, Color> richTextBoxAppend)
    {
        [GeneratedRegex("tdchain|sachain|td|sa|mbist|hardip", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("^(?!Mbist)M([a-zA-Z]){1}([a-zA-Z0-9]){2}([a-zA-Z0-9]){1,2}$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        private static readonly Regex _regTest = MyRegex();

        private readonly BinCutScriptMain _binCutScriptMain = binCutScriptMain;
        private readonly Action<string, Color> _richTextBoxAppend = richTextBoxAppend;
        private readonly string _tempFolder = binCutScriptMain.TempFolder;
        protected string BinCutFilePath = binCutScriptMain.BinCutFilePath;
        protected string BinCutPostFilePath = binCutScriptMain.BinCutPostFilePath;
        protected string ModeSeqenceFilePath = binCutScriptMain.ModeSequenceFilePath;
        protected string TestProgramFilePath = binCutScriptMain.TestProgramFilePath;
        protected string TestPlanFilePath = binCutScriptMain.TestPlanFilePath;
        protected string PowerBinningPath = binCutScriptMain.PowerBinning;
        protected string IdsDistributionFilePath = binCutScriptMain.IdsDistributionFilePath;
        protected string PatternDashBoardFilePath = binCutScriptMain.PatDashFilePath;
        protected string DataLogFolder = string.Empty;
        protected string OutPutFolder = binCutScriptMain.OutPutFolder;
        protected string OutputFile = string.Empty;
        public Job Job = binCutScriptMain.Job;
        private ExcelWorkbook? _testplanWorkbook;

        private void AddGoodBinsFromTempFolder()
        {
            if (string.IsNullOrEmpty(_tempFolder))
            {
                return;
            }
            _richTextBoxAppend("Reading Bintables for Good Bin List ...", Color.Blue);
            List<string> bintables = [.. Directory.GetFiles(_tempFolder, "Bin_Table*.txt")];
            var goodBins = new List<string>();
            foreach (string bintable in bintables)
            {
                var binSheetReader = new ReadBinTableSheet();
                BinTableSheet binSheet = binSheetReader.GetSheet(bintable);
                List<BinTableRow> rows = [.. binSheet.Rows.Where(x => !x.IsBackup)];
                IEnumerable<string> currentGoodBins = rows
                    .FindAll(x => x.Result == "Pass" && x.Bin.Length != 0)
                    .Select(x => x.Bin)
                    .Distinct();
                goodBins.AddRange(currentGoodBins);
            }
            BinCutData.GoodBins.AddRange(goodBins.Distinct());
        }

        private void ReadNewHarvestDssc(bool isCsharp)
        {

            var candidateSheetNames = new List<string>
            {
                "Dig_Src_instructions",
                $"Dig_Src_instructions_{Job.JobName}"
            };

            if (isCsharp)
            {
                candidateSheetNames.Add("UF_DigSrc_HarvMappingTable");
                candidateSheetNames.Add("UF_DigSrc_HARVMappingTable");
            }

            ExcelWorksheet? digSrcInstructionSheet = candidateSheetNames
                .Select(name => IgxlSheetReaderHelpers.GetExcelWorksheet(TestProgramFilePath, name))
                .FirstOrDefault(sheet => sheet != null);
            if (digSrcInstructionSheet == null)
            {
                return;
            }

            _richTextBoxAppend("Reading Dig_Src_instrctions ...", Color.Blue);
            BinCutData.DigSrcInstructionSheet = new DigSrcInstrctionReaderForHarvestSourceCodeCompare().ReadSheet(digSrcInstructionSheet);
        }

        private void LoadPowerPins(BinningTables binningTables, BinCutFlowTables binCutFlowTables)
        {
            BinCutData.ModeVsCorePowerName = binningTables.First().GetModeVsPowerName();
            BinCutData.PowerPins = binCutFlowTables.GetTitle(Job.JobType);
            foreach (string powerPin in BinCutData.PowerPins)
            {
                if (BinCutData.ModeVsCorePowerName.ContainsValue(powerPin))
                {
                    BinCutConfig.PowerType.TryAdd(powerPin, EnumPowerType.Core_Power);
                }
                else
                {
                    BinCutConfig.PowerType.TryAdd(powerPin, EnumPowerType.Others);
                }
            }
        }

        private void LoadIdsDistributionTable(bool isCsharp)
        {
            if (string.IsNullOrEmpty(IdsDistributionFilePath))
            {
                string idsfile = Path.Combine(_tempFolder, "IDS_Distribution.TXT");
                string isdfileWithJob = Path.Combine(_tempFolder, "IDS_Distribution_" + Job.JobName + ".TXT");

                if (File.Exists(idsfile))
                {
                    IdsDistributionFilePath = idsfile;
                }

                if (File.Exists(isdfileWithJob))
                {
                    IdsDistributionFilePath = isdfileWithJob;
                }
            }

            if (!string.IsNullOrEmpty(IdsDistributionFilePath))
            {
                _richTextBoxAppend("Reading IDS Distribution Table ...", Color.Blue);
                BinCutData.IdsdistributionTable = IdsDistributionReader.Read(IdsDistributionFilePath);
            }
            else if (isCsharp)
            {
                string idsfile = Path.Combine(_tempFolder, "bincut_starting_eqn_table.TXT");
                IdsDistributionFilePath = idsfile;
                if (!string.IsNullOrEmpty(IdsDistributionFilePath) && File.Exists(idsfile))
                {
                    _richTextBoxAppend("Reading Start EQN Table ...", Color.Blue);
                    BinCutData.IdsdistributionTable = IdsDistributionReader.ReadStartEqn(IdsDistributionFilePath);
                }
            }
        }

        private void SepvmFusing(ExcelWorkbook excelWorkbook)
        {
            var sepvmFusingSheetList = new List<ExcelWorksheet>();
            sepvmFusingSheetList.AddRange(BinCutReadMainHelpers1.GetSheetFromTempFolder("SEPVM_fusing*.txt", _tempFolder));
            if (sepvmFusingSheetList.Count == 0 && _testplanWorkbook != null)
            {
                sepvmFusingSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(_testplanWorkbook.Worksheets, "SEPVM_fusing"));
            }
            else if (sepvmFusingSheetList.Count == 0 && excelWorkbook != null)
            {
                sepvmFusingSheetList.AddRange(BinCutReadMainHelpers.GetBinCutExcelSheetList(excelWorkbook.Worksheets, "SEPVM_fusing"));
            }

            if (sepvmFusingSheetList.Count != 0)
            {
                _richTextBoxAppend("Reading SEPVM_fusing ...", Color.Blue);
                //ExcelWorksheet sheet = sepvmFusingSheetList[0];
                //var sepvmfusingReader = new SepvmfusingSheetReader();
                //SepvmfusingSheet sepvmfusingSheet = sepvmfusingReader.ReadSheet(sheet);
                //BinCutData.SepvmfusingSheet = sepvmfusingSheet;
                BinCutData.IsSepvmfusing = true;
            }
        }

        public bool ReadBinCutData(bool isCsharp, string? iniFileName = null, ExcelWorksheet? excelWorksheet = null)
        {
            try
            {
                Initialize();
                BinCutReadMainHelpers.LoadBinCutConfig();

                TestProgramReaderHelpers.ReadTestProgram(isCsharp, _richTextBoxAppend, TestProgramFilePath, _tempFolder);
                AddGoodBinsFromTempFolder();
                HarvestHelpers.ReadHarvestDssc(TestProgramFilePath, Job, _richTextBoxAppend);

                // Read HarvPinGrpTable sheet from temp folder into BinCutData
                //if (File.Exists(Path.Combine(_tempFolder, "HarvestPinFlag_Table.txt")))
                //{
                //    string harvPinGroup = Path.Combine(_tempFolder, "HarvestPinFlag_Table.txt");
                //    BinCutData.HarvPinGroupsheet = new HarvPinGroupTableReader.HarvPinGroup().Read(harvPinGroup);
                //}

                ReadNewHarvestDssc(isCsharp);

                List<string> performanceModeList = [];
                List<TestSettingData>? testSettingSheetsList = null;
                // Reading test plan
                if (!string.IsNullOrEmpty(TestPlanFilePath))
                {
                    var jobs = new List<string> { Job.JobType.ToString() };
                    if (jobs == null || jobs.Count == 0)
                    {
                        testSettingSheetsList = new TestSettingReader(null).ReadFlow(_binCutScriptMain.ProjectName, _testplanWorkbook!, []);
                    }
                    else
                    {
                        testSettingSheetsList = new TestSettingReader(null).ReadFlowByAllJobs(_binCutScriptMain.ProjectName, _testplanWorkbook!, jobs, null);
                    }
                    DcCategoryInfos categoryInfoList = [];
                    foreach (TestSettingData testSetting in testSettingSheetsList)
                    {
                        foreach (DcCategoryName category in testSetting.DcCategorys)
                        {
                            if (!categoryInfoList.Exists(s => s.CategoryName.EqualsIgnoreCase(category.CategoryName)))
                            {
                                categoryInfoList.Add(category.DcCategoryInfo);
                            }
                        }
                    }

                    IEnumerable<DcCategoryInfo> temp = categoryInfoList.Where(x => _regTest.IsMatch(x.Test) && MyRegex1().IsMatch(x.PmodePatternVdip));
                    performanceModeList = [.. temp.Select(x => x.PmodePatternVdip)];

                    _richTextBoxAppend("Reading Test Plan ...", Color.Blue);
                    var testPlanExcelPackage = new ExcelPackage(new FileInfo(TestPlanFilePath));
                    _testplanWorkbook = testPlanExcelPackage.Workbook;
                    BinCutData.HasBinCutInstance = testPlanExcelPackage.Workbook.Worksheets.Any(sheet =>
                        sheet.Name.StartsWithIgnoreCase("Bincut_Instance")
                        || sheet.Name.StartsWithIgnoreCase("Instance_Bincut_")
                    );
                }

                // Export binning sheets from BinCut test plan
                if (!string.IsNullOrEmpty(BinCutFilePath))
                {
                    _richTextBoxAppend("Exporting binning sheets from BinCut test plan ...", Color.Blue);
                    Dictionary<string, string> modes = BinCutReadMainHelpers1.GetAllPerformanceModeDic(performanceModeList);
                    BinCutService.NonIgxlSheetProcess(modes, isCsharp, new ExcelPackage(new FileInfo(BinCutFilePath)).Workbook, _tempFolder, true);
                }

                BinCutReadMainHelpers.ReadPatternDashboard(PatternDashBoardFilePath);

                GetBinningSpecificData(out BinningTables binningTables, out List<string> binningTitleList, out Dictionary<string, List<string>> domainDic);

                // Try get BinCut workbook
                ExcelPackage? binCutExcelPackage = BinCutReadMainHelpers.GetBinCutExcelPackage(isCsharp, BinCutFilePath, _tempFolder);
                ExcelWorkbook? binCutWorkbook = binCutExcelPackage?.Workbook;

                _richTextBoxAppend("Reading Flow Sheet ...", Color.Blue);
                BinCutFlowTables binCutFlowTables = BinCutReadMainHelpers.GetBinCutFlowTables(binCutWorkbook!, isCsharp, binningTitleList, domainDic, Job, _tempFolder);

                _richTextBoxAppend("Reading Post Flow Sheet ...", Color.Blue);
                BinCutReadMainHelpers.LoadBinCutPostFlowSheet(isCsharp, binningTitleList, domainDic, BinCutPostFilePath, _tempFolder);

                // Load Mode Sequence Sheet
                if (!string.IsNullOrEmpty(ModeSeqenceFilePath))
                {
                    _richTextBoxAppend("Reading Mode Sequence Sheet ...", Color.Blue);
                }

                LoadPowerPins(binningTables, binCutFlowTables);
                LoadIdsDistributionTable(isCsharp);

                // SelSRAM
                if (TestProgramFilePath.Length != 0 && (BinCutData.SelsrmMappingSheet == null || BinCutData.SelsrmMappingSheet.Rows.Count == 0))
                {
                    _richTextBoxAppend("Reading SELSRM_Mapping_Table ...", Color.Blue);
                    var selsramMappingTableSheets = new List<ExcelWorksheet>();
                    selsramMappingTableSheets.AddRange(BinCutReadMainHelpers1.GetSheetFromTempFolder("SELSRM_Mapping_Table.txt", _tempFolder));
                    selsramMappingTableSheets.AddRange(BinCutReadMainHelpers1.GetSheetFromTempFolder("SELSRM_Mapping_Table_CP.txt", _tempFolder));
                    if (selsramMappingTableSheets.Count != 0)
                    {
                        var selsrmMappingSheetReader = new SelsrmMappingSheetReader();
                        BinCutData.SelsrmMappingSheet = selsrmMappingSheetReader.ReadSheet(selsramMappingTableSheets.First()) ?? new SelsrmMappingSheet();
                    }
                }

                ReadPowerBinningMain.StartPowerBinning(PowerBinningPath, _tempFolder, _testplanWorkbook!, _richTextBoxAppend);

                _richTextBoxAppend("Reading new Pwrbin_Seq ...", Color.Blue);
                ReadPowerBinningMain.ReadNewPowerBinningSheet(binCutWorkbook!, _testplanWorkbook!, PowerBinningPath, _tempFolder, Job);

                SepvmFusing(binCutWorkbook!);

                _richTextBoxAppend("Reading NV/VRS table ...", Color.Blue);
                BinCutReadMainHelpers.GenNvVrsTable(testSettingSheetsList!, Job, _binCutScriptMain.IsVoltageByProgram, _tempFolder);
            }
            catch (Exception e)
            {
                if (_binCutScriptMain.CmdMode)
                {
                    _richTextBoxAppend(e.Message + Environment.NewLine + e.StackTrace, Color.Red);
                    throw;
                }
                if (e.Message.StartsWithIgnoreCase("Autogen"))
                {
                    ErrorMessageBox.Show(e.Message, "");
                }
                else
                {
                    ErrorMessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "");
                }

                return false;
            }

            return true;
        }

        private void GetBinningSpecificData(out BinningTables binningTables, out List<string> binningTitleList, out Dictionary<string, List<string>> domainDic)
        {
            // Get Binning specific data
            binningTables = BinCutReadMainHelpers1.GetVddBinningDefAndOtherRail(_tempFolder, _richTextBoxAppend, Job);
            binningTitleList = [];
            domainDic = [];
            if (binningTables.Any())
            {
                binningTitleList = binningTables.GetBinningTitleList();
                domainDic = binningTables.GetDomainDic();
            }
        }

        private void Initialize()
        {

            BinCutData.Initialize();
            if (Directory.Exists(_tempFolder))
            {
                Directory.Delete(_tempFolder, true);
            }

            Directory.CreateDirectory(_tempFolder);
        }
    }
}
