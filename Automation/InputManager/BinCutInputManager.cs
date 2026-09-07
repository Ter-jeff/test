using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenAc;
using Automation.GenerateIgxl.Basic.Business.GenAc.AcGenerator.Business;
using Automation.GenerateIgxl.Basic.Business.GenAc.AcInput.BassData;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;
using Automation.GenerateIgxl.HardIp;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.TpUpdate.HardIPEnableWordsUpdate;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using LogLib.Static;

using OfficeOpenXml;


using TestPlanLib.Basic;
using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Checker;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.FlowNew;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;
using TestPlanLib.Static;

namespace Automation.InputManager
{

    public class BinCutInputManager : InputManagerBase<BinCutInputData>
    {
        private readonly List<string> _hipSyntaxList = new List<string>();
        private readonly Dictionary<string, List<EnableWordTableRow>> _hipPatMappingTable = new Dictionary<string, List<EnableWordTableRow>>();
        private List<UfDigSrcSheet> _ufDigSrcSheet;
        private Dictionary<string, EnableWordTable> _ttrTable;

        public BinCutInputManager(ExcelWorkbook excelWorkbook) : base(excelWorkbook)
        {
        }

        public override BinCutInputData Read()
        {
            var binCutInputData = new BinCutInputData { Config = SettingStatic.ScanConfig };

            #region Read UF_DigSrc_XXX
            if (TestPlanStatic.UfDigSrcSheets != null && TestPlanStatic.UfDigSrcSheets.Any())
            {
                _ufDigSrcSheet = TestPlanStatic.UfDigSrcSheets;
            }
            #endregion

            #region Read Binning
            ReadBinning(binCutInputData);
            #endregion

            var binningTitleList = new List<string>();
            var domainDic = new Dictionary<string, List<string>>();
            if (binCutInputData.BinningTables.Any())
            {
                binningTitleList = binCutInputData.BinningTables.GetBinningTitleList();
                domainDic = binCutInputData.BinningTables.GetDomainDic();
            }

            #region Read binCut flow sheet
            ReadBinCutFlowSheet(binCutInputData, binningTitleList, domainDic);
            #endregion

            #region Read binCut shadow flow sheet
            ReadBinCutShadowFlowSheet(binCutInputData, binningTitleList, domainDic);
            #endregion

            #region Read binCut flow Post sheet
            ReadBinCutPostFlowSheet(binCutInputData, binningTitleList, domainDic);
            #endregion

            #region Read BinCut Mode Sequence
            ReadBinCutModeSequence(binCutInputData);
            #endregion

            #region Read bincut_instance_post sheet
            ReadBinCutInstancePost(binCutInputData);
            #endregion

            #region Read BinCutOrder
            ReadBinCutOrder(binCutInputData);
            #endregion

            #region Read Bincut_Instance is split multiple sheet for TD, BIST and BIRA
            ReadBinCutInstance(binCutInputData);
            #endregion

            #region Read BinCut Instance NamingRule
            if (binCutInputData.BinCutInstanceSheets.Count != 0)
            {
                binCutInputData.BinCutInstanceNamingSheet = SettingStatic.BinCutInstanceNamingSheet;
            }
            #endregion

            #region Read HardIp
            if (binCutInputData.BinCutInstanceSheets.Count != 0 || binCutInputData.BinCutInstancePostSheets.Count != 0)
            {
                AddBinCutInstanceFromTtrTable(_hipSyntaxList, binCutInputData.BinCutInstanceSheets, binCutInputData.BinCutInstancePostSheets, ref binCutInputData);
            }

            #endregion

            #region Read RtosCategoryConfigWorkbook
            ReadRtosCategoryConfigWorkbook();
            #endregion



            return binCutInputData;
        }

        private void ReadBinning(BinCutInputData binCutInputData)
        {
            if (EpWorkbook.BinCutWorkbook == null)
            {
                return;
            }

            Response.Report("Reading Binning File ...", percentage: 20);
            bool isCs = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder);
            var vddBinList = new List<ExcelWorksheet>();
            if (EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.Binning] != null)
            {
                vddBinList.Add(EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.Binning]);
            }

            if (EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BinningBinX] != null)
            {
                vddBinList.Add(EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BinningBinX]);
            }

            if (EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BinningBinY] != null)
            {
                vddBinList.Add(EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BinningBinY]);
            }

            if (vddBinList.Count == 0)
            {
                return;
            }

            binCutInputData.BinningTables = new BinningTables();
            for (int index = 0; index < vddBinList.Count; index++)
            {
                BinningTable binningTable = BinningTableReader.Read(vddBinList[index], "All");
                binningTable.AddToErrorReport();
                binCutInputData.BinningTables.Add(binningTable);
                if (isCs)
                {
                    IdsDistributionTable idsDistributionTable = IdsDistributionReader.ReadStartEqnSheet(LocalSpecs.BinCutFileName, "All");
                    binCutInputData.IdsdistributionTable.AllIdsPowers.AddRange(idsDistributionTable.AllIdsPowers);
                }
            }
            foreach (string fileName in LocalSpecs.BinCutShadowFileNames)
            {
                string shadowStage = Path.GetFileNameWithoutExtension(fileName).Split('_').Last();
                var package = new ExcelPackage(new FileInfo(fileName));
                ExcelWorksheet excelWorksheet = package.Workbook.Worksheets[NeededSheets.Binning];
                if (excelWorksheet != null)
                {
                    BinningTable binningTable = BinningTableReader.Read(excelWorksheet, shadowStage);
                    binCutInputData.BinningTables.Add(binningTable);
                }
                if (isCs)
                {
                    IdsDistributionTable idsDistributionTable = IdsDistributionReader.ReadStartEqnSheet(fileName, shadowStage);
                    binCutInputData.IdsdistributionTable.AllIdsPowers.AddRange(idsDistributionTable.AllIdsPowers);
                }
            }
            foreach (BinningTable vddBinTable in binCutInputData.BinningTables)
            {
                vddBinTable.AddToErrorReport();
            }
        }

        private void ReadBinCutFlowSheet(BinCutInputData binCutInputData, List<string> binningTitleList, Dictionary<string, List<string>> domainDic)
        {
            if (EpWorkbook.BinCutWorkbook == null)
            {
                return;
            }

            ExcelWorksheet binCutWorksheet = EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BcFlow] ??
                                             EpWorkbook.BinCutWorkbook.Worksheets[NeededSheets.BcAte];
            if (binCutWorksheet == null)
            {
                return;
            }

            Response.Report("Reading BinCut Flow File ...", percentage: 20);
            var binCutSheetReader = new BinCutFlowSheetReader(binningTitleList, domainDic);
            BinCutFlowTables binCutFlowTables = binCutSheetReader.ReadSheet(binCutWorksheet);
            binCutFlowTables.Check(binCutInputData.BinningTables);
            binCutInputData.BinCutFlowTables = binCutFlowTables;
            if (EpWorkbook.BinCutModeSeqWorkbook == null)
            {
                _hipSyntaxList.AddRange(binCutInputData.BinCutFlowTables.SelectMany(x => x.Rows).Select(x => x.SpiRtos).Where(x => Regex.IsMatch(x, @"^M([a-zA-Z]){1}([a-zA-Z0-9]){2}([a-zA-Z0-9]){1,2} \w+ (ELB|ILB)$")).Distinct());
            }
            foreach (BinCutFlowTable binCutFlowTable in binCutFlowTables)
            {
                binCutFlowTable.AddToErrorReport();
            }
        }

        private void ReadBinCutShadowFlowSheet(BinCutInputData binCutInputData, List<string> binningTitleList, Dictionary<string, List<string>> domainDic)
        {
            if (LocalSpecs.BinCutShadowFileNames == null)
            {
                return;
            }

            foreach (string fileName in LocalSpecs.BinCutShadowFileNames)
            {
                var package = new ExcelPackage(new FileInfo(fileName));
                ExcelWorksheet excelWorksheet = package.Workbook.Worksheets[NeededSheets.BcAte];
                if (excelWorksheet != null)
                {
                    Response.Report("Reading BinCut Flow Shadow File ...", percentage: 20);
                    var binCutSheetReader = new BinCutFlowSheetReader(binningTitleList, domainDic);
                    BinCutFlowTables binCutShadowFlowTables = binCutSheetReader.ReadSheet(excelWorksheet);
                    binCutShadowFlowTables.Check(binCutInputData.BinningTables);
                    binCutInputData.BinCutFlowShadowTables.AddRange(binCutShadowFlowTables);
                    if (EpWorkbook.BinCutModeSeqWorkbook == null)
                    {
                        _hipSyntaxList.AddRange(binCutInputData.BinCutFlowShadowTables.SelectMany(x => x.Rows).Select(x => x.SpiRtos).Where(x => Regex.IsMatch(x, @"^M([a-zA-Z]){1}([a-zA-Z0-9]){2}([a-zA-Z0-9]){1,2} \w+ (ELB|ILB)$")).Distinct());
                    }
                    foreach (BinCutFlowTable binCutShadowFlowTable in binCutShadowFlowTables)
                    {
                        binCutShadowFlowTable.AddToErrorReport();
                    }
                }
            }
        }

        private void ReadBinCutPostFlowSheet(BinCutInputData binCutInputData, List<string> binningTitleList, Dictionary<string, List<string>> domainDic)
        {
            if (EpWorkbook.BinCutPostWorkbook == null)
            {
                return;
            }

            Response.Report("Reading BinCut Post Flow File ...", percentage: 20);
            foreach (ExcelWorksheet sheet in EpWorkbook.BinCutPostWorkbook.Worksheets)
            {
                if (sheet.Name.StartsWith("Flow_PostBinCut", StringComparison.CurrentCultureIgnoreCase))
                {
                    var binCutPostSheetReader = new BinCutFlowSheetReader(binningTitleList, domainDic);
                    BinCutFlowTables postFlowTables = binCutPostSheetReader.ReadSheet(sheet);
                    postFlowTables.Check(binCutInputData.BinningTables);
                    binCutInputData.BinCutPostFlowSheets.Add(postFlowTables);
                    foreach (BinCutFlowTable postFlowTable in postFlowTables)
                    {
                        postFlowTable.AddToErrorReport();
                    }
                }
            }
            if (binCutInputData.BinCutPostFlowSheets.Count > 1)
            {
                ValidateBinCutPost.CheckFlowSheetsPmode(binCutInputData.BinCutPostFlowSheets);
                foreach (BinCutFlowTables binCutFlowTables in binCutInputData.BinCutPostFlowSheets)
                {
                    foreach (BinCutFlowTable binCutFlowTable in binCutFlowTables)
                    {
                        binCutFlowTable.AddToErrorReport();
                    }
                }
            }
        }

        private void ReadBinCutModeSequence(BinCutInputData binCutInputData)
        {
            if (EpWorkbook.BinCutModeSeqWorkbook == null)
            {
                return;
            }

            foreach (ExcelWorksheet sheet in EpWorkbook.BinCutModeSeqWorkbook.Worksheets)
            {
                ReadBinCutModeSequenceSheet(binCutInputData, sheet);
            }
            _hipSyntaxList.AddRange(binCutInputData.NewBinCutFlowTables.SelectMany(x => x.Rows).SelectMany(x => x.SubFlows).Select(x => x.Split(':').Last()).Where(y => Regex.IsMatch(y, @"^M([a-zA-Z]){1}([a-zA-Z0-9]){2}([a-zA-Z0-9]){1,2} (((\w+) (ELB|ILB))|(ELB|ILB))")).Distinct());
            _hipSyntaxList.AddRange(binCutInputData.NewPostBinCutFlowTables.SelectMany(x => x.Rows).SelectMany(x => x.SubFlows).Select(x => x.Split(':').Last()).Where(y => Regex.IsMatch(y, @"^M([a-zA-Z]){1}([a-zA-Z0-9]){2}([a-zA-Z0-9]){1,2} (((\w+) (ELB|ILB))|(ELB|ILB))")).Distinct());
        }

        private void ReadBinCutModeSequenceSheet(BinCutInputData binCutInputData, ExcelWorksheet sheet)
        {
            if ((sheet.Name.StartsWith(NeededSheets.BinCutModeSequence, StringComparison.CurrentCultureIgnoreCase) ||
                sheet.Name.StartsWith(NeededSheets.BcFlow, StringComparison.CurrentCultureIgnoreCase)) && binCutInputData.BinCutFlowTables.Any())
            {
                Response.Report("Reading BinCut Mode Sequence File ...", percentage: 20);
                bool result = true;
                var binCutSheetReader = new NewBincutFlowSheetReaderV2(ref result);
                NewBinCutFlowTables binCutFlowTables = binCutSheetReader.ReadSheet(sheet);
                if (!result)
                {
                    var binCutSheetReaderOld = new BinCutFlowSheetReaderNew();
                    binCutFlowTables = binCutSheetReaderOld.ReadSheet(sheet);
                }
                binCutFlowTables.Check(binCutInputData.BinningTables);
                binCutInputData.NewBinCutFlowTables = binCutFlowTables;
                foreach (NewBinCutFlowTable binCutFlowTable in binCutFlowTables)
                {
                    List<NewBinCutFlowSheetRow> postBinCutFlowTableRows = binCutFlowTable.Rows.FindAll(x => x.TableType.Equals(EnumBinCutTableType.Post));
                    if (postBinCutFlowTableRows != null && postBinCutFlowTableRows.Any())
                    {
                        NewBinCutFlowTable postBinCutFlowTable = binCutFlowTable.Copy();
                        postBinCutFlowTable.Rows.Clear();
                        postBinCutFlowTable.Rows.AddRange(postBinCutFlowTableRows);
                        binCutInputData.NewPostBinCutFlowTables.Add(postBinCutFlowTable);
                    }
                }
            }
            else if (sheet.Name.StartsWith(NeededSheets.PostBinCutModeSeq, StringComparison.CurrentCultureIgnoreCase) && binCutInputData.BinCutPostFlowSheets.Any())
            {
                Response.Report("Reading Post BinCut Mode Sequence File ...", percentage: 20);
                var binCutSheetReader = new BinCutFlowSheetReaderNew();
                NewBinCutFlowTables binCutFlowTables = binCutSheetReader.ReadSheet(sheet);
                binCutFlowTables.Check(binCutInputData.BinningTables);
                binCutInputData.NewPostBinCutFlowTables = binCutFlowTables;
            }
            else if (
                sheet.Name.StartsWith("BV", StringComparison.CurrentCultureIgnoreCase) ||
                sheet.Name.StartsWith("HBV", StringComparison.CurrentCultureIgnoreCase))
            {
                Response.Report("Reading BinCut Mode Sequence File ...", percentage: 20);
                bool result = true;
                var binCutSheetReader = new NewBincutFlowSheetReaderV3(ref result);
                NewBinCutFlowTables binCutFlowTables = binCutSheetReader.ReadSheet(sheet);
                binCutFlowTables.Check(binCutInputData.BinningTables);
                binCutInputData.NewBinCutFlowTables.AddRange(binCutFlowTables);
                binCutInputData.NewBinCutFlowTables.JobList = binCutSheetReader.ReadJobCount();
            }
            else if (sheet.Name.StartsWith("PBC", StringComparison.CurrentCultureIgnoreCase))
            {
                Response.Report("Reading Post BinCut Mode Sequence File ...", percentage: 20);
                bool result = true;
                var binCutSheetReader = new NewBincutFlowSheetReaderV3(ref result);
                NewBinCutFlowTables binCutFlowTables = binCutSheetReader.ReadSheet(sheet);
                binCutFlowTables.Check(binCutInputData.BinningTables);
                binCutInputData.NewPostBinCutFlowTables = binCutFlowTables;
                binCutInputData.NewPostBinCutFlowTables.JobList = binCutSheetReader.ReadJobCount();
            }
        }

        private void ReadBinCutInstancePost(BinCutInputData binCutInputData)
        {
            if (EpWorkbook.BinCutPostWorkbook == null)
            {
                return;
            }

            Dictionary<string, PatternData> patternDatas = AcTSetCategoryMapSingleton.Instance().PatternList;
            Dictionary<string, int> timeSets = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
            var bcPostSheets = EpWorkbook.TestPlanWorkbook.Worksheets.Where(sheet => sheet.Name.StartsWith("BinCut_Instance_Post", StringComparison.OrdinalIgnoreCase) || sheet.Name.StartsWith("Instance_BinCut_post", StringComparison.OrdinalIgnoreCase) || sheet.Name.StartsWith("Instance_post_BinCut", StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (ExcelWorksheet bcPostSheet in bcPostSheets)
            {
                var binCutInstanceSheetReader = new BinCutInstanceSheetReader();
                BinCutInstanceSheet binCutInstanceSheet = binCutInstanceSheetReader.ReadSheet(bcPostSheet);
                new BinCutInstanceSheetChecker().WorkFlow(binCutInstanceSheet, binCutInputData.BinCutPostFlowSheets, binCutInputData.NewPostBinCutFlowTables, patternDatas, timeSets, _ufDigSrcSheet);
                if (binCutInstanceSheet != null)
                {
                    binCutInputData.BinCutInstancePostSheets.Add(binCutInstanceSheet);
                    binCutInstanceSheet.AddToErrorReport();
                }

                foreach (BinCutFlowTables binCutFlowTables in binCutInputData.BinCutPostFlowSheets)
                {
                    foreach (BinCutFlowTable binCutFlowTable in binCutFlowTables)
                    {
                        binCutFlowTable.AddToErrorReport();
                    }
                }
            }
            CheckFlowNames(binCutInputData.BinCutPostFlowSheets, binCutInputData.BinCutInstancePostSheets, binCutInputData.NewPostBinCutFlowTables);
        }

        private void ReadBinCutOrder(BinCutInputData binCutInputData)
        {
            ExcelWorksheet workSheet = EpWorkbook.TestPlanWorkbook.Worksheets["BinCutOrder"];
            if (workSheet == null && EpWorkbook.BinCutWorkbook != null)
            {
                workSheet = EpWorkbook.BinCutWorkbook.Worksheets["BinCutOrder"];
            }

            if (workSheet != null && binCutInputData.BinCutFlowTables.Any())
            {
                Response.Report("Reading BinCutOrder ...", percentage: 20);
                var binCutOrderReader = new BinCutOrderReader(binCutInputData.BinCutFlowTables);
                binCutInputData.BinCutOrderSheet = binCutOrderReader.ReadSheet(workSheet);
                binCutInputData.BinCutOrderSheet.AddToErrorReport();
            }
        }

        private void ReadBinCutInstance(BinCutInputData binCutInputData)
        {
            Dictionary<string, PatternData> patternData = AcTSetCategoryMapSingleton.Instance().PatternList;
            Dictionary<string, int> timeSet = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
            foreach (BinCutInstanceSheet binCutInstanceSheet in TestPlanStatic.BinCutInstanceSheets)
            {
                new BinCutInstanceSheetChecker().WorkFlow(binCutInstanceSheet, new BinCutFlowSheets { binCutInputData.BinCutFlowTables }, binCutInputData.NewBinCutFlowTables, patternData, timeSet, _ufDigSrcSheet);
                if (binCutInstanceSheet != null)
                {
                    binCutInputData.BinCutInstanceSheets.Add(binCutInstanceSheet);
                }

                binCutInstanceSheet.AddToErrorReport();
            }
            CheckFlowNames(new BinCutFlowSheets { binCutInputData.BinCutFlowTables }, binCutInputData.BinCutInstanceSheets, binCutInputData.NewBinCutFlowTables);

            if (AcTSetCategoryMapSingleton.Instance().TimeSetBlock2Categories == null ||
               AcTSetCategoryMapSingleton.Instance().TimeSetBlock2Categories.Count == 0)
            {
                var patList = patternData.Values.ToList();
                var timeSetGenerator = new TimeSetGenerator();
                TimeSetSheets multiTimeSetSheets = timeSetGenerator.GenerateFlow(patList, LocalSpecs.TimeSetFolder, FolderStructure.DirTimings);
                AcInputSheet acInputSheet = new BasicInitial().InitalAcSymbols();
                var acGenerator = new AcGenerator(acInputSheet);
                AcSpecSheet acSpecSheet = acGenerator.GenerateFlow(multiTimeSetSheets, patList);
                TestProgram.IgxlWorkBk.AddAcSpecSheet(FolderStructure.DirAcSpec, acSpecSheet);
            }
        }

        private void ReadRtosCategoryConfigWorkbook()
        {
            if (SettingStatic.RtosCategoryConfigWorkbook != null)
            {
                return;
            }

            string rtosConfigPath =
                Path.Combine(LocalSpecs.SettingFolder, "Settings", "Basic", $"RtosCategory_{LocalSpecs.CurrentProject}.xlsx");
            if (!string.IsNullOrEmpty(LocalSpecs.SettingFiles.RtosCategoryConfig))
            {
                rtosConfigPath = LocalSpecs.SettingFiles.RtosCategoryConfig;
            }

            if (File.Exists(rtosConfigPath))
            {
                Response.Report("Reading Rtos Category Config ...", percentage: 20);
                var inputExcel = new ExcelPackage(new FileInfo(rtosConfigPath));
                SettingStatic.RtosCategoryConfigWorkbook = inputExcel.Workbook;
            }
        }



        internal void AddBinCutInstanceFromTtrTable(List<string> hipSyntaxList, List<BinCutInstanceSheet> instanceSheets, List<BinCutInstanceSheet> postInstanceSheets,
            ref BinCutInputData binCutInputData)
        {
            if (!string.IsNullOrEmpty(LocalSpecs.TtrSummaryFileName) &&
                BlockStatus.GetAutomationBlockStatus(BlockStatus.HardIp).Down && (LocalSpecs.IsModuleIncluded("IDS") || LocalSpecs.IsModuleIncluded(BlockStatus.HardIp)))
            {
                _ttrTable = new TtrToolReader().ReadEnableWordTables(LocalSpecs.TtrSummaryFileName);
                if (GetHipPatBySyntax(hipSyntaxList).Any())
                {
                    var hipInstSheet = new BinCutInstanceSheet("TTRTable");
                    var hipInstSheetPost = new BinCutInstanceSheet("TTRTable_Post");
                    foreach (KeyValuePair<string, List<EnableWordTableRow>> hardIpPattern in _hipPatMappingTable)
                    {
                        List<EnableWordTableRow> allPatterns = hardIpPattern.Value;
                        string syntax = hardIpPattern.Key;
                        string bcType;
                        if (syntax.Contains("HBV") && syntax.Contains("PBC"))
                        {
                            bcType = "HBV_Chk";
                        }
                        else if (syntax.Contains("PBC"))
                        {
                            bcType = "BV_Chk";
                        }
                        else if (syntax.Contains("HBV"))
                        {
                            bcType = "HBV";
                        }
                        else
                        {
                            bcType = "BV";
                        }

                        foreach (EnableWordTableRow pattern in allPatterns)
                        {
                            var row = new BinCutInstanceRow("TTRTable") { FlowNameOri = syntax, Type = BincutInstanceType.Hardip };
                            if (Regex.IsMatch(pattern.Name, @"^Multiple_", RegexOptions.IgnoreCase))
                            {
                                pattern.Name = Regex.Replace(pattern.Name, @"^Multiple_", "", RegexOptions.IgnoreCase);
                            }
                            row.PatternList.Add(pattern.Name);
                            row.PayloadList.Add(pattern.Name);
                            row.PatWithIndex.Add(0, pattern.Name);
                            row.JobTestStage = string.Join(",", FoundAvailableJobs(pattern, bcType));
                            switch (bcType)
                            {
                                case "BV_Chk":
                                    row.DCcategory = "LV";
                                    break;
                                case "HBV_Chk":
                                    row.DCcategory = "HV";
                                    break;
                                case "HBV":
                                    row.DCcategory = "HV";
                                    break;
                                default:
                                    row.DCcategory = "LV";
                                    break;
                            }
                            if (!string.IsNullOrEmpty(row.JobTestStage))
                            {
                                if (bcType == "BV_Chk" || bcType == "HBV_Chk")
                                {
                                    hipInstSheetPost.Rows.Add(row);
                                }
                                else
                                {
                                    hipInstSheet.Rows.Add(row);
                                }
                            }
                        }
                    }
                    if (hipInstSheet.Rows.Any())
                    {
                        instanceSheets.Add(hipInstSheet);
                    }

                    if (hipInstSheetPost.Rows.Any())
                    {
                        postInstanceSheets.Add(hipInstSheetPost);
                    }
                }
            }
            Response.Report("Reading HardIp ...", percentage: 20);
            var patterns = instanceSheets.SelectMany(x => x.Rows).Where(x => x.Type == BincutInstanceType.Hardip).SelectMany(x => x.PatternList).ToList();
            var postPatterns = postInstanceSheets.SelectMany(x => x.Rows).Where(x => x.Type == BincutInstanceType.Hardip).SelectMany(x => x.PatternList).ToList();
            if (postPatterns.Any())
            {
                patterns.AddRange(postPatterns);
            }
            if (patterns.Any())
            {
                binCutInputData.HardIpPatterns = GetHardipInstanceRow(patterns);
            }
        }

        public Dictionary<string, HardIpSheet> GetHardipInstanceRow(List<string> usedPatterns)
        {
            var hardIpParaData = new HardIpParaData(EnumBlock.HardIp);
            var hardIpMain = new HardIpMain();
            return hardIpMain.ReadHardipSheets(hardIpParaData, usedPatterns);
        }

        private List<string> FoundAvailableJobs(EnableWordTableRow pattern, string bcType)
        {
            var jobs = new List<string>();
            if (!pattern.EnableWords.TryGetValue(bcType, out List<string> value))
            {
                return jobs;
            }

            jobs.AddRange(value.Select(x => x.Replace("Prod_", "")));
            return jobs;
        }

        private List<EnableWordTableRow> GetHipPatBySyntax(List<string> syntaxList)
        {
            var foundPattern = new List<EnableWordTableRow>();
            var allPattern = new List<EnableWordTableRow>();
            syntaxList = syntaxList.Select(x => x.ToUpper()).ToList();
            foreach (KeyValuePair<string, EnableWordTable> enableWordTable in _ttrTable)
            {
                allPattern.AddRange(enableWordTable.Value.Rows);
            }
            foreach (string syntax in syntaxList)
            {
                string mode = syntax.Split(' ').First();
                string type = syntax.EndsWith("HBV", StringComparison.CurrentCultureIgnoreCase) || syntax.EndsWith("PBC", StringComparison.CurrentCultureIgnoreCase) ? syntax.Split(' ')[1] : syntax.Split(' ').Last();

                List<EnableWordTableRow> allPatterns = allPattern.FindAll(x => x.Name.Contains(mode) && x.Name.Contains(type));

                if (!_hipPatMappingTable.ContainsKey(syntax))
                {
                    _hipPatMappingTable.Add(syntax, allPatterns);
                }

                foreach (EnableWordTableRow pattern in allPatterns.Where(pattern => !foundPattern.Contains(pattern)))
                {
                    foundPattern.Add(pattern);
                }
            }
            return foundPattern;
        }

        private void CheckFlowNames(BinCutFlowSheets binCutFlowSheets, List<BinCutInstanceSheet> binCutInstanceSheets, NewBinCutFlowTables newBinCutFlowTables = null)
        {
            var binCutInstance = binCutInstanceSheets.SelectMany(x => x.Rows).Select(x => x.FlowName).Distinct().ToList();
            List<string> flowNames = newBinCutFlowTables != null && newBinCutFlowTables.Any()
                ? newBinCutFlowTables.GetFlowNames()
                : binCutFlowSheets.GetFlowNames();

            foreach (string item in flowNames.Distinct().ToList())
            {
                CheckFlowName(item.Trim(), binCutInstance);
            }
        }

        private void CheckFlowName(string flowName, List<string> flowNameInInstance)
        {
            if (string.IsNullOrEmpty(flowName) || flowName.ContainsIgnoreCase("_DONOTEST") || flowName.ContainsIgnoreCase("_DONOTTEST") || flowName.ContainsIgnoreCase("ELB") || flowName.ContainsIgnoreCase("ILB") || flowName.ToUpper().StartsWith("FLOW_NWIRE") || flowName.ToUpper().StartsWith("RELAY_"))
            {
                return;
            }

            if (!flowNameInInstance.Exists(x => x.Equals(flowName, StringComparison.CurrentCultureIgnoreCase)))
            {
                string path = LocalSpecs.BinCutModeSeqFileName.Replace(Directory.GetCurrentDirectory(), "");
                ErrorReportManager.AddError(BinCutErrorType.E_FlowName_01, path, 0, 0, "The flowName \"" + flowName + "\" is missing in the all BinCut_Instance sheet !!!", new string[] { flowName });
            }
        }
    }
}
