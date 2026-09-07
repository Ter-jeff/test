using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.PostAction.TempMon.Enums;
using Automation.GenerateIgxl.Scan.Harvest.Flow;
using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Atpg;
using Automation.Utility.Atpg.Data;
using Automation.Utility.Basic;

using CommonLib.Enums;
using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc;
using TestPlanLib.BinNumber;
using TestPlanLib.Concurrent;
using TestPlanLib.Harvest;
using TestPlanLib.Singleton;
using TestPlanLib.Utility;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Scan.Harvest
{
    public class ScanNonBinCutInstanceMain
    {
        protected virtual string Module => "Scan";
        internal static readonly Regex _regex3 = new Regex(@"\w+:\w+\((?<Flag>\w+)\)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        protected readonly ScanConfig Config;
        protected DataTable PayloadTypeTable;

        private const string BinTableNonBinCut = "Bin_Table_Non_Bincut";
        protected List<HarvestingTruthTableSheet> TruthTables;
        private readonly MultiTestSettingSheetsSingleton _multiTestSettingSheetsSingleton;
        private readonly List<string> _performanceModeList;
        private ExcelPackage _missingPinsReport;
        private readonly List<string> _harvestCheckedPattern;
        private readonly BinCutFlowTable _equationTable;
        protected readonly HashSet<string> _controlBinOutAndInstanceFlags;

        public ScanNonBinCutInstanceMain(ScanConfig config)
        {
            Config = config;
            PayloadTypeTable = config.PayloadType;
            if (EpWorkbook.EquationVoltages != null && EpWorkbook.EquationVoltages.Worksheets["EquationVoltages"] != null)
            {
                _equationTable = new EquationVoltageReader().ReadSheet(EpWorkbook.EquationVoltages.Worksheets["EquationVoltages"]).First();
            }
            TruthTables = TestPlanStatic.HarvestingTruthTableSheets;
            _multiTestSettingSheetsSingleton = MultiTestSettingSheetsSingleton.Instance();
            _performanceModeList = MultiTestSettingSheetsSingleton.Instance().PerformanceModeList;
            _harvestCheckedPattern = new List<string>();
            _controlBinOutAndInstanceFlags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        public virtual void WorkFlow()
        {
            var patSets = new List<PatSet>();
            var flows = new List<SubFlowSheet>();
            var instanceHarvestSheet = new InstanceSheet("TestInst_Harvest");
            var instanceSheets = new List<InstanceSheet>();
            var binTableRows = new List<BinTableRow>();
            var binTableSheet = new BinTableSheet(BinTableNonBinCut);

            BinCutInstanceNamingSheet binCutInstanceNamingSheet = BinCutInstanceNamingSheet();
            var binCutInstanceSheets = GetBinCutInstanceSheets().ToList();

            var instSheetPreProcess = new InstSheetPreProcess(Config);
            BinCutFinalInstanceRows binCutFinalInstanceRows = instSheetPreProcess.InitialInstance(binCutInstanceSheets, binCutInstanceNamingSheet);
            binCutFinalInstanceRows = binCutFinalInstanceRows.RePatSetNameDuplicateRows();

            instanceSheets.Add(new InstanceSheet("TestInst_CpuScan"));
            instanceSheets.Add(new InstanceSheet("TestInst_SocScan"));
            instanceSheets.Add(new InstanceSheet("TestInst_GfxScan"));
            instanceSheets.Add(new InstanceSheet("TestInst_Non_Bincut"));
            instanceSheets.ForEach(sheet => sheet.SourceInfo.Block = nameof(EnumBlock.Scan));
            instanceHarvestSheet.SourceInfo.Block = nameof(EnumBlock.Harvest);
            AcSpecSheet acSpecSheet = TestProgram.IgxlWorkBk.GetAcSpecsSheet();
            if (acSpecSheet != null)
            {
                new BinCutAcSpecsWriter().GenAcSpecs(binCutFinalInstanceRows, acSpecSheet);
            }

            var harvestInstanceRows = new List<InstanceRow>();
            GenInstance(binCutFinalInstanceRows, ref instanceSheets);
            flows.AddRange(GenFlow(binCutFinalInstanceRows, TruthTables));

            binTableRows.AddRange(GetBinTableRows(binCutFinalInstanceRows));

            List<PatSet> patSetSheet = GenPatSets(binCutFinalInstanceRows);
            patSetSheet = AddCommandAndFlagInPatSet(patSetSheet, binCutFinalInstanceRows);
            patSets.AddRange(patSetSheet);

            if (TestPlanStatic.MappingCoreTable != null)
            {
                List<BinCutFinalInstanceRow> ssnInstance = binCutFinalInstanceRows.FindAll(x => x.IsSsn());
                flows.Add(MappingCoreTable.CreateCoreMaskSubFlow());
                IEnumerable<string> existSsnType = ssnInstance.Select(x => $"{x.Domain}_{x.Block}".ToUpper()).ToList().Distinct();
                foreach (string ssnType in existSsnType)
                {
                    var ssnBinTableRow = new BinTableRow { Name = $"Bin_{ssnType}_SSN_INIT_Fail", ItemList = "F_SSN_INIT_Fail", Op = "AND", Items = new List<string> { "T" } };
                    BinNumResult ssnBinInfo = BinNumberSingleton.Instance.GetBinInfo("Scan", ssnType.Split('_').First(), ssnType.Split('_').Last(), ssnBinTableRow);
                    ssnBinTableRow.Sort = ssnBinInfo.SoftBin.ToString("G15");
                    ssnBinTableRow.Bin = ssnBinInfo.BinNumInfo.HardBin.ToString("G15");
                    ssnBinTableRow.Result = ssnBinInfo.BinNumInfo.Status;
                    binTableRows.Add(ssnBinTableRow);
                    Function functionSsn = TestProgram.VbtFunctionLib.GetFunctionByName("MaskSsnCores", "scan");
                    if (functionSsn.Type == ".NET")
                    {
                        instanceHarvestSheet.AddRow(new InstanceRow { TestName = "SSN_CoreMask", VbtType = ".NET", VbtName = functionSsn.FullFunctionName });
                    }
                    else
                    {
                        instanceHarvestSheet.AddRow(new InstanceRow { TestName = "SSN_CoreMask", VbtType = "VBT", VbtName = "SSN_CoreMask" });
                    }
                }
            }

            if (TruthTables.Any())
            {
                AddHarvestArtifacts(flows, harvestInstanceRows, binTableRows, instanceHarvestSheet, instanceSheets);
            }

            #region Add into igxl
            List<BinTableRow> binTableGroup = RemoveDuplicateBinTableRows(binTableRows);
            foreach (BinTableRow binTableRow in binTableGroup)
            {
                binTableSheet.AddRow(binTableRow);
            }

            TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirNonBinCut, binTableSheet);

            List<string> flags = GenAllFailFlagToMainFlow(binTableSheet);
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(flags, "Scan", FolderStructure.DirMain);
            if (TruthTables.Any())
            {
                foreach (HarvestingTruthTableSheet truthTable in TruthTables)
                {
                    var list = truthTable.Rows.First().Fusings.SelectMany(x => x.GetAddressFlags()).ToList();
                    TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(list, "Harvest", FolderStructure.DirMain);
                }
            }

            SetPatSetSheet(patSets);

            foreach (SubFlowSheet flow in flows)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirNonBinCut, flow);
            }

            foreach (InstanceSheet instanceSheet in instanceSheets)
            {
                if (instanceSheet.Rows.Any())
                {
                    TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirNonBinCut, instanceSheet);
                }
            }
            #endregion
            AddControlBinFlagsToMainInitEnableWd();
        }

        protected void AddControlBinFlagsToMainInitEnableWd()
        {
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(_controlBinOutAndInstanceFlags, $"{Module}InstBinOut", FolderStructure.DirMain);
        }

        private void AddHarvestArtifacts(List<SubFlowSheet> flows, List<InstanceRow> harvestInstanceRows, List<BinTableRow> binTableRows, InstanceSheet instanceHarvestSheet, List<InstanceSheet> instanceSheets)
        {
            if (TestPlanStatic.HarvestingFlagInitSheet != null)
            {
                TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(TestPlanStatic.HarvestingFlagInitSheet.GetFlags(), "HarvestInit", FolderStructure.DirMain);
            }
            if (TestPlanStatic.FlagOperationSheets.Any())
            {
                foreach (FlagOperationSheet sheet in TestPlanStatic.FlagOperationSheets)
                {
                    flows.Add(sheet.CreateSubFlow());
                    TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(sheet.GetAllFlags(), "Harvest", FolderStructure.DirMain);
                }
            }
            if (TestPlanStatic.AteStrSummarySheet != null)
            {
                flows.AddRange(TestPlanStatic.AteStrSummarySheet.CreateFlows());
                harvestInstanceRows.Add(CreateAteSummaryInstance());
                IEnumerable<string> repairFlags = TruthTables.SelectMany(x => x.GetRepairFlagList()).Distinct();
                TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(TestPlanStatic.AteStrSummarySheet.GetAllFlags().Except(repairFlags, StringComparer.OrdinalIgnoreCase), "AteStrSummary - TestName", FolderStructure.DirMain);
                TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(TestPlanStatic.AteStrSummarySheet.GetAllFlagFromValue().Except(repairFlags, StringComparer.OrdinalIgnoreCase), "AteStrSummary - Value", FolderStructure.DirMain);
            }

            if (TestPlanStatic.DomainFlagsSheets != null)
            {
                var allHarvFlagsFromDomain = new List<string>();
                foreach (DomainFlagsSheet sheet in TestPlanStatic.DomainFlagsSheets)
                {
                    allHarvFlagsFromDomain.AddRange(sheet.GetAllFlags());
                }
                TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(allHarvFlagsFromDomain, "Harvest", FolderStructure.DirMain);
            }

            if (TestPlanStatic.DigitalFlagsSheet != null)
            {
                TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(TestPlanStatic.DigitalFlagsSheet.Rows.Select(x => x.FlagName).Distinct(), "DigitalFlags", FolderStructure.DirMain);
            }

            foreach (HarvestingTruthTableSheet truthTable in TruthTables)
            {
                var harvestMain = new ScanHarvestMain(truthTable);
                harvestMain.GenHarvesting(flows, harvestInstanceRows);
                BinTableRow binTable = harvestMain.GenHarvestOtherBinTableRow();
                if (!binTableRows.Exists(x => x.Name.Equals(binTable.Name)))
                {
                    binTableRows.Add(binTable);
                }

                BinTableRow binTableReadBack = harvestMain.GenHarvestReadBackBinTableRow();
                if (!binTableRows.Exists(x => x.Name.Equals(binTableReadBack.Name)))
                {
                    binTableRows.Add(binTableReadBack);
                }

                truthTable.AddToErrorReport();
            }
            var harvestPostCheck = new ScanHarvestMain(null);
            if (harvestPostCheck.HasHarvestPostCheckFlow)
            {
                binTableRows.Add(harvestPostCheck.GenHarvestPostCheckBinTable());
                flows.Add(harvestPostCheck.GenHarvPostCheckSubFlow(harvestInstanceRows));
            }
            flows.Add(new ScanHarvestMain(null).GenHarvReadSubFlow(harvestInstanceRows));
            instanceHarvestSheet.AddRows(harvestInstanceRows);
            instanceSheets.Add(instanceHarvestSheet);
        }

        private InstanceRow CreateAteSummaryInstance()
        {
            var instanceRow = new InstanceRow { TestName = "ATE_STR_Summary" };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("HarvestingPrintFlagsStatus", "scan", true);
            if (function.Type == ".NET")
            {
                instanceRow.VbtType = ".NET";
                instanceRow.VbtName = function.FullFunctionName;
                function.SetParamValue("flagsToPrint", string.Join(",", TestPlanStatic.AteStrSummarySheet.GetAllFlags()));
                instanceRow.ArgList = function.Parameters;
                instanceRow.Args = function.ArgList;
            }

            return instanceRow;
        }

        internal List<BinTableRow> RemoveDuplicateBinTableRows(List<BinTableRow> binTableRows)
        {
            List<BinTableRow> rows = binTableRows.Select(x => x.Copy()).ToList();
            foreach (BinTableRow duplicateRow in binTableRows.Select(y => binTableRows.FindAll(x => y.Name.Equals(x.Name) && y.ItemList.Split(new[] { " ", "," }, StringSplitOptions.RemoveEmptyEntries).SequenceEqual(x.ItemList.Split(new[] { " ", "," }, StringSplitOptions.RemoveEmptyEntries)))).SelectMany(duplicateRows => duplicateRows))
            {
                rows.Remove(duplicateRow);
            }
            return rows;
        }

        internal BinCutInstanceNamingSheet BinCutInstanceNamingSheet()
        {
            return SettingStatic.BinCutInstanceNamingSheet;
        }

        public static List<string> GetBinCutInstanceSheets(List<string> sheetList)
        {
            var sheetNames = new List<string>();
            foreach (string sheet in sheetList)
            {
                if (sheet.StartsWith("BinCut_Instance", StringComparison.CurrentCultureIgnoreCase) || sheet.StartsWith("Instance_BinCut_", StringComparison.CurrentCultureIgnoreCase) || sheet.StartsWith("Instance_Post_BinCut", StringComparison.CurrentCultureIgnoreCase) || sheet.ContainsIgnoreCase("clock") || sheet.ContainsIgnoreCase("evs") || sheet.ContainsIgnoreCase("htol") || sheet.ContainsIgnoreCase("cpm"))
                {
                    continue;
                }

                if (sheet.StartsWith("Instance_", StringComparison.CurrentCultureIgnoreCase))
                {
                    sheetNames.Add(sheet);
                }
            }
            return sheetNames;
        }

        private List<BinCutInstanceSheet> GetBinCutInstanceSheets()
        {
            var binCutInstanceSheets = new List<BinCutInstanceSheet>();
            Dictionary<string, PatternData> dictionary = AcTSetCategoryMapSingleton.Instance().PatternList;
            Dictionary<string, int> timeSets = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
            var ufDigSrcSheet = new List<UfDigSrcSheet>();
            if (TestPlanStatic.UfDigSrcSheets != null && TestPlanStatic.UfDigSrcSheets.Any())
            {
                ufDigSrcSheet = TestPlanStatic.UfDigSrcSheets;
            }
            foreach (BinCutInstanceSheet binCutInstanceSheet in TestPlanStatic.ScanInstanceSheets)
            {
                new NonBinCutInstanceSheetChecker().WorkFlow(binCutInstanceSheet, dictionary, timeSets, ufDigSrcSheet);
                if (binCutInstanceSheet != null)
                {
                    binCutInstanceSheets.Add(binCutInstanceSheet);
                    binCutInstanceSheet.AddToErrorReport();
                    if (binCutInstanceSheet.HasPatternPinGroups())
                    {
                        if (_missingPinsReport == null)
                        {
                            InitMissPinReport();
                        }

                        AtpgService.CheckHarvestGroupPins(
                            binCutInstanceSheet,
                            _missingPinsReport,
                            _harvestCheckedPattern);
                        SaveMissReport();
                    }
                }
            }
            return binCutInstanceSheets;
        }

        private void InitMissPinReport()
        {
            string outputPath = Path.Combine(LocalSpecs.TarFolder, "MissingPinReport.xlsx");
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            _missingPinsReport = new ExcelPackage(new FileInfo(outputPath));
            ExcelWorksheet extraSheet = _missingPinsReport.Workbook.AddSheet("Extra harvest pins");
            ExcelWorksheet missSheet = _missingPinsReport.Workbook.AddSheet("Miss harvest pins");
            var titles = new List<string> { "Pattern", "Pins" };
            extraSheet.Cells[1, 1].PrintExcelRow(titles.ToArray());
            missSheet.Cells[1, 1].PrintExcelRow(titles.ToArray());
        }

        private void SaveMissReport()
        {
            _missingPinsReport.Save();
        }

        internal List<string> GenAllFailFlagToMainFlow(BinTableSheet binTableSheet)
        {
            var failList = new List<string>();
            foreach (BinTableRow binRow in binTableSheet.Rows)
            {
                if (binRow.Items.Count > 1)
                {
                    continue; //skip HLV item
                }

                if (Regex.IsMatch(binRow.ItemList, "_Group", RegexOptions.IgnoreCase))
                {
                    if (!failList.Exists(x => x.Equals(binRow.ItemList, StringComparison.OrdinalIgnoreCase)))
                    {
                        failList.Add(binRow.ItemList);
                    }
                }
            }
            return failList;
        }

        internal virtual List<PatSet> GenPatSets(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            var patSets = new List<PatSet>();
            foreach (BinCutFinalInstanceRow binCutFinalInstanceRow in binCutFinalInstanceRows)
            {
                if (binCutFinalInstanceRow.PatternList.Count < 2)
                {
                    continue;
                }

                var patSet = new PatSet
                {
                    PatSetName = binCutFinalInstanceRow.PatSetName,
                    Domain = binCutFinalInstanceRow.Domain
                };

                foreach (string pattern in binCutFinalInstanceRow.PatternList)
                {
                    var patSetRow = new PatSetRow { File = pattern };
                    if (binCutFinalInstanceRow.IsBurstNonBinCutInstance())
                    {
                        patSetRow.Burst = "Yes";
                    }
                    else
                    {
                        patSetRow.Burst = "No";
                    }

                    patSetRow.Comment = "SheetName: " + binCutFinalInstanceRow.BinCutInstanceRow.SheetName + ", Row: " + binCutFinalInstanceRow.BinCutInstanceRow.RowNum;
                    patSet.AddRow(patSetRow);
                }
                if (!patSets.Exists(x => x.PatSetName.Equals(patSet.PatSetName)))
                {
                    patSets.Add(patSet);
                }
            }
            return patSets;
        }

        internal virtual void SetPatSetSheet(List<PatSet> patSets)
        {
            if (patSets.Count > 0)
            {
                var patSetNonBinCut = new PatSetSheet("PatSets_Non_Bincut");
                var patSetSoc = new PatSetSheet("PatSets_SocScan");
                var patSetCpu = new PatSetSheet("PatSets_CpuScan");
                var patSetGfx = new PatSetSheet("PatSets_GfxScan");
                foreach (IGrouping<string, PatSet> patSet in patSets.GroupBy(x => x.Domain))
                {
                    switch (patSet.Key)
                    {
                        case "Soc":
                            {
                                patSetSoc.AddRows(patSet.ToList());
                                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirNonBinCut, patSetSoc);
                                break;
                            }
                        case "Cpu":
                            {
                                patSetCpu.AddRows(patSet.ToList());
                                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirNonBinCut, patSetCpu);
                                break;
                            }
                        case "Gfx":
                            {
                                patSetGfx.AddRows(patSet.ToList());
                                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirNonBinCut, patSetGfx);
                                break;
                            }
                        default:
                            {
                                patSetNonBinCut.AddRows(patSet.ToList());
                                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirNonBinCut, patSetNonBinCut);
                                break;
                            }
                    }
                }
            }
        }

        internal virtual List<FlowRow> GetBinFlowRows(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            HashSet<FlowRow> binFlowRowsSet = new HashSet<FlowRow>(AtpgBinOutService.GetBinOutItemsComparer<FlowRow>());
            binFlowRowsSet.AddRange(GetBinOutDatas(binCutFinalInstanceRows).Where(x => !x.IsByPassBinOut).SelectMany(AtpgBinOutService.GetBinFlowRow));
            return binFlowRowsSet.ToList();
        }

        internal virtual List<BinTableRow> GetBinTableRows(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            HashSet<BinTableRow> binTableRowsSet = new HashSet<BinTableRow>(AtpgBinOutService.GetBinOutItemsComparer<BinTableRow>());
            binTableRowsSet.AddRange(GetBinOutDatas(binCutFinalInstanceRows).Where(x => !x.IsByPassBinOut).SelectMany(x => AtpgBinOutService.GetBinTableRow(x, Module)));
            return binTableRowsSet.ToList();
        }

        internal virtual List<AtpgBinOutData> GetBinOutDatas(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            List<AtpgBinOutData> resultsList = new List<AtpgBinOutData>();
            List<string> allHarvestFlags = TruthTables?.FirstOrDefault()?.GetHarvAllFlagList() ?? new List<string>();
            foreach (BinCutFinalInstanceRow binCutFinalInstanceRow in binCutFinalInstanceRows)
            {
                string pattern = binCutFinalInstanceRow.PayloadList == null || !binCutFinalInstanceRow.PayloadList.Any() ? binCutFinalInstanceRow.PatternList.First() : binCutFinalInstanceRow.PayloadList.First();
                string payloadType = GetPayloadType(pattern);
                List<string> failActionFlags = binCutFinalInstanceRow.GetBinOutFlags(payloadType, binCutFinalInstanceRow.PerformanceMode);
                List<string> failflags = failActionFlags.Except(allHarvestFlags, StringComparer.OrdinalIgnoreCase).ToList();

                resultsList.AddRange(AtpgBinOutService.GetAtpgBinOutDatas(failflags, binCutFinalInstanceRow));
            }
            return resultsList;
        }

        private void GenInstance(List<BinCutFinalInstanceRow> binCutFinalInstanceRows, ref List<InstanceSheet> instanceSheets)
        {
            var allInstanceRows = new InstanceRows();
            var allInstanceNames = new HashSet<string>();
            var concurrentFlow = new ConcurrentFlowSheet("Concurrent_Flow");
            if (EpWorkbook.TestPlanWorkbook.Worksheets["Concurrent_Flow"] != null)
            {
                concurrentFlow = TestPlanStatic.ConcurrentFlowSheet;
            }
            IEnumerable<IGrouping<string, BinCutFinalInstanceRow>> groups = binCutFinalInstanceRows.GroupBy(x => BinCutInstanceRowUtility.GetDomain(x.BinCutInstanceRow));
            var existFlowHeaderFooter = new List<string>();
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                InstanceSheet instanceSheet = instanceSheets.Find(x => x.Name.Equals(GetInstanceSheetNameByDomain(group.Key)));
                foreach (BinCutFinalInstanceRow row in group)
                {
                    var instanceRow = new InstanceRow();
                    string selector = row.GetVoltageType();
                    instanceRow.TestName = row.GetParameter();
                    string oriSheetName = string.IsNullOrEmpty(row.BinCutInstanceRow.SubFlow) ? row.BinCutInstanceRow.FlowName : row.BinCutInstanceRow.SubFlow;

                    if (allInstanceNames.Contains(instanceRow.TestName.ToUpper()))
                    {
                        row.IsDuplicateName = true;
                        instanceRow.TestName = row.GetParameter();
                    }
                    instanceRow.TimeSets = row.GetTimeSetVersion(row.PatternList);
                    instanceRow.DcCategory = GetDcCategory(row);
                    instanceRow.DcSelector = GetDcSelector(selector);
                    instanceRow.AcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow.AcSpec) ? GetAcCategory(row.BinCutInstanceRow, instanceRow.TimeSets) : row.BinCutInstanceRow.AcSpec;
                    instanceRow.AcSelector = "Typ";
                    instanceRow.PinLevels = concurrentFlow.Rows.Exists(x => x.Subflows.Exists(y => y.ToUpper().Equals(row.BinCutInstanceRow.FlowName.ToUpper())))
                        ? GenerateLevelConcurrent(row.BinCutInstanceRow.FlowName, instanceRow.TimeSets, concurrentFlow, row.BinCutInstanceRow.Levels)
                            : GenerateLevel(instanceRow.DcCategory, row.Block, row.BinCutInstanceRow.Levels);
                    instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
                    if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
                    {
                        instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
                        instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
                    }
                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameFunctionalT, "scan", true);
                    if (function.Type == ".NET")
                    {
                        AtpgService.GenerateCSharpInstanceRow(row, instanceRow, function);
                    }
                    else
                    {
                        function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameFunctionalT, "scan");
                        GenerateVbtInstanceRow(row, instanceRow, function);
                    }
                    TempMonService.TrySetTempMon(LocalSpecs.TempMonDatas, row.BinCutInstanceRow.TempMon, instanceRow.TestName, EnumType.Instance);
                    allInstanceRows.Add(instanceRow);
                    allInstanceNames.Add(instanceRow.TestName.ToUpper());
                    instanceSheet.AddRow(instanceRow);
                    if (!existFlowHeaderFooter.Exists(x => x.Equals(oriSheetName, StringComparison.OrdinalIgnoreCase)))
                    {
                        allInstanceRows.AddHeaderFooter(oriSheetName);
                        instanceSheet.AddHeaderFooter(oriSheetName);
                        existFlowHeaderFooter.Add(oriSheetName);
                    }
                }
            }
        }

        private void GenerateVbtInstanceRow(BinCutFinalInstanceRow row, InstanceRow instanceRow, Function vbtFunctionBase)
        {
            string dsscPat = "";
            if (row.PatternList.Any())
            {
                foreach (string pat in row.PatternList)
                {
                    if (Regex.IsMatch(pat, @"\w*DSSC\w*", RegexOptions.IgnoreCase))
                    {
                        dsscPat = pat;
                    }
                }
            }
            if (row.GetDcCategory().Contains("_EQN") && LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                vbtFunctionBase.SetParamValue("ApplyVoltageFromBinCut", row.BinCutInstanceRow.DCcategory);
                List<string> items = row.PatSetName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                items.Add("EQN");
                row.PatSetName = string.Join("_", items.Where(x => !string.IsNullOrEmpty(x)));
                instanceRow.TestName = row.GetParameter();
            }
            vbtFunctionBase.SetParamValue("Patterns", row.PatternList.Count == 1 ? row.PatternList[0] : row.PatSetName);
            if (row.IsBurstNonBinCutInstance())
            {
                vbtFunctionBase.SetParamValue("ResultMode", "1");
            }
            else
            {
                vbtFunctionBase.SetParamValue("ResultMode", "0");
            }

            vbtFunctionBase.SetParamValue("RelayMode", "1");
            vbtFunctionBase.SetParamValue("UserFunction", row.BinCutInstanceRow.UserFunction);
            if (Regex.IsMatch(row.BinCutInstanceRow.MultiFstpEnable, "TRUE", RegexOptions.IgnoreCase))
            {
                vbtFunctionBase.SetParamValue("MultiFSTP_Enable", "TRUE");
            }

            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.PatternPinGroup))
            {
                string payloadType = GetPayloadType(row.BinCutInstanceRow.PayloadList.First());
                vbtFunctionBase.SetParamValue("Harv_FailFlag", row.BinCutInstanceRow.PatternPinGroup);
                List<string> otherFlags = OtherFlags(row.BinCutInstanceRow);
                vbtFunctionBase.SetParamValue("HarvestPinGrpOtherFail", otherFlags.Any() ? string.Join(",", otherFlags) : row.GenFlag(payloadType, row.PerformanceMode));
            }

            vbtFunctionBase.SetParamValue("Harv_Condition", row.BinCutInstanceRow.PatternPinGroup);

            if (!string.IsNullOrEmpty(dsscPat))
            {
                string sendPinName = "JTAG_TDI";
                if (LocalSpecs.HardIpInfos != null)
                {
                    HardIpInfo target = LocalSpecs.HardIpInfos.GetHardIpInfo(dsscPat);
                    if (target != null && !string.IsNullOrEmpty(target.SendPinName))
                    {
                        sendPinName = target.SendPinName;
                    }
                }
                vbtFunctionBase.SetParamValue("DigSource", "Test_AutoSwitch:" + sendPinName.ToUpper());

            }

            instanceRow.VbtType = "VBT";
            instanceRow.VbtName = vbtFunctionBase.FullFunctionName;
            instanceRow.ArgList = vbtFunctionBase.Parameters;
            instanceRow.Args = vbtFunctionBase.ArgList;
        }

        public List<string> OtherFlags(BinCutInstanceRow binCutInstanceRow)
        {
            var pinFlags = binCutInstanceRow.PatternPinGroup.Split(';').Select(x => x.Split('(').Last().Trim(')')).ToList();
            string[] binoutFlags = binCutInstanceRow.PinGroupBinoutFlag.Split([','], StringSplitOptions.RemoveEmptyEntries);
            List<string> otherflags = [];
            foreach (string binoutFlag in binoutFlags)
            {
                if (!pinFlags.Contains(binoutFlag))
                {
                    otherflags.Add(binoutFlag);
                }
            }
            return otherflags;
        }

        internal string GetInstanceSheetNameByDomain(string domain)
        {
            if (domain.Contains("Soc"))
            {
                return "TestInst_SocScan";
            }
            else if (domain.Contains("Cpu"))
            {
                return "TestInst_CpuScan";
            }
            else if (domain.Contains("Gfx"))
            {
                return "TestInst_GfxScan";
            }
            else
            {
                return "TestInst_Non_Bincut";
            }
        }

        protected virtual string GetPerformanceModes(BinCutFinalInstanceRow row)
        {
            string mode;
            if (string.IsNullOrEmpty(row.PerformanceMode))
            {
                mode = AtpgService.GetPerformanceMode(row.InitList, _performanceModeList);
            }
            else
            {
                mode = row.PerformanceMode;
            }

            if (_performanceModeList.Exists(p => p.Equals(mode, StringComparison.OrdinalIgnoreCase)))
            {
                return mode.ToUpper();
            }

            if (!string.IsNullOrEmpty(mode))
            {
                return mode;
            }

            return "";
        }

        internal string GetDcSelector(string selectorName)
        {
            string selector = "Typ";
            switch (selectorName)
            {
                case "HV":
                    selector = "Max";
                    break;
                case "LV":
                    selector = "Min";
                    break;
                case "NV":
                    selector = "Typ";
                    break;
            }
            return selector;
        }

        internal virtual string GetDcCategory(BinCutFinalInstanceRow binCutFinalInstanceRow)
        {
            string mode = GetPerformanceModes(binCutFinalInstanceRow);
            if (!string.IsNullOrEmpty(binCutFinalInstanceRow.BinCutInstanceRow.DCcategory))
            {
                string dcCategoryUserDefine = GetDcCategory(binCutFinalInstanceRow.BinCutInstanceRow.DCcategory);
                if (dcCategoryUserDefine.Contains("_EQN"))
                {
                    if (_equationTable != null)
                    {
                        if (_equationTable.Rows.Exists(x => x.PerformanceMode.Equals(binCutFinalInstanceRow.BinCutInstanceRow.DCcategory, StringComparison.OrdinalIgnoreCase)))
                        {
                            dcCategoryUserDefine = GetDcCategory(_equationTable.Rows.Find(x => x.PerformanceMode.Equals(binCutFinalInstanceRow.BinCutInstanceRow.DCcategory, StringComparison.OrdinalIgnoreCase)).AllOther);
                        }
                    }
                    dcCategoryUserDefine = dcCategoryUserDefine.Replace("_EQN", "");
                }
                if (_multiTestSettingSheetsSingleton != null)
                {
                    if (_multiTestSettingSheetsSingleton.DcCategoryInfos.Exists(s => s.CategoryName.Equals(dcCategoryUserDefine, StringComparison.OrdinalIgnoreCase)))
                    {
                        return dcCategoryUserDefine;
                    }
                }
                else
                {
                    return dcCategoryUserDefine;
                }
            }

            string domain = binCutFinalInstanceRow.Domain;
            string type = GetPayloadType(binCutFinalInstanceRow.PayloadList != null && binCutFinalInstanceRow.PayloadList.Any() ? binCutFinalInstanceRow.PayloadList[0] : null);
            string dcCategory = _multiTestSettingSheetsSingleton.FindScanCategoryName(type, domain, mode, binCutFinalInstanceRow.PayloadList, out EnumMessageLevel _, out string _, domain);
            if (string.IsNullOrEmpty(dcCategory))
            {
                dcCategory = type + "_" + domain + "_X_" + mode;
            }

            return dcCategory;
        }

        protected string GetAcCategory(BinCutInstanceRow instanceDataRow, string timeSet)
        {
            return AtpgService.GenerateGetAcCategory(instanceDataRow, timeSet);
        }

        protected string GenerateLevelConcurrent(string flowName, string timeSet, ConcurrentFlowSheet concurrentFlow, string userDefinedLevel)
        {
            if (!string.IsNullOrEmpty(userDefinedLevel))
            {
                return userDefinedLevel;
            }

            if (AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets.Find(x => x.Name.Equals(timeSet)) == null)
            {
                return "TBD(ConcurrentLevelError)";
            }

            string timeDomain =
                AcTSetCategoryMapSingleton.Instance()
                    .MultiTimeSetSheets.Find(x => x.Name.Equals(timeSet))
                    .TimeDomain;
            return "Levels_Con_SC_" + timeDomain;
        }

        internal virtual string GenerateLevel(string dcCategory = "", string block = "", string userDefinedLevel = "")
        {
            if (!string.IsNullOrEmpty(userDefinedLevel))
            {
                return userDefinedLevel;
            }

            if (dcCategory.ContainsIgnoreCase("EVS"))
            {
                return "Levels_EVS_Scan";
            }

            if (dcCategory.ContainsIgnoreCase("BIST") || dcCategory.ContainsIgnoreCase("BIRA"))
            {
                return "Levels_Mbist";
            }

            if (dcCategory.ContainsIgnoreCase("BINCUT") && _multiTestSettingSheetsSingleton.DcCategoryInfos.IsSplitDcSpecs(LocalSpecs.Options.IsSplitDcSpecs))
            {
                return "Levels_Bincut";
            }

            return "Levels_Scan";
        }

        internal string GetDcCategory(string category)
        {
            string dcCategory = Regex.Replace(category, @"\s*_*(HV|LV|NV)$", "", RegexOptions.IgnoreCase);
            return dcCategory;
        }

        protected virtual List<SubFlowSheet> GenFlow(List<BinCutFinalInstanceRow> binCutFinalInstanceRows, List<HarvestingTruthTableSheet> truthTables = null)
        {
            var flows = new List<SubFlowSheet>();
            IEnumerable<IGrouping<string, BinCutFinalInstanceRow>> groups = binCutFinalInstanceRows.GroupBy(x => x.GetSubFlowName());
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                string sheetName = group.Key;
                string testPlanSheet = group.FirstOrDefault()?.BinCutInstanceRow.SheetName;
                var flow = new SubFlowSheet(sheetName, testPlanSheet + ":" + sheetName);
                if (group.Select(x => x.BinCutInstanceRow).ToList().Where(x => !string.IsNullOrEmpty(x.JobTestStage))
                    .Select(x => x.JobTestStage.Replace("!", "")).Distinct().Count() == 1)
                {
                    string jobName = group.Select(x => x.BinCutInstanceRow).Select(x => x.JobTestStage).First();
                    if (!jobName.Contains(","))
                    {
                        flow.JobNames = new List<string> { jobName };
                    }
                }
                flow.AddStartRows();
                flow.AddRows(GenFlowBody(group));
                flow.AddRows(GenFlowBinTable(group.ToList(), flow, truthTables));
                flow.AddEndRows();
                flows.Add(flow);
            }
            return flows;
        }

        public virtual List<FlowRow> GenFlowBinTable(List<BinCutFinalInstanceRow> group, SubFlowSheet subFlowSheet = null, List<HarvestingTruthTableSheet> truthTables = null)
        {
            if (subFlowSheet != null && truthTables != null)
            {
                GenHarvest(group, subFlowSheet, truthTables);
            }
            var flowRows = new List<FlowRow>();
            List<FlowRow> binFlowRows = GetBinFlowRows(group).OrderByDescending(x => x.BinTableFlagCount).ToList();
            foreach (FlowRow binFlowRow in binFlowRows)
            {
                (FlowRow ifRow, FlowRow endIfRow) = GetIfControlBinRow(binFlowRow);
                flowRows.Add(ifRow);
                flowRows.Add(binFlowRow);
                flowRows.Add(endIfRow);
            }
            return flowRows;
        }

        public (FlowRow, FlowRow) GetIfControlBinRow(FlowRow binFlowRow)
        {
            FlowRow ifRow = new FlowRow() { Opcode = OpCode.If, Parameter = $"Control{binFlowRow.Parameter}", };
            FlowRow endIfRow = new FlowRow() { Opcode = OpCode.EndIf };
            return (ifRow, endIfRow);
        }

        private void GenHarvest(List<BinCutFinalInstanceRow> group, SubFlowSheet subFlowSheet, List<HarvestingTruthTableSheet> truthTables)
        {
            var failFlags = group.ToList().SelectMany(x => GetAllFlags(x.BinCutInstanceRow.PatternPinGroup)).Distinct().ToList();
            bool hasHarvItems = group.ToList().Exists(x => !string.IsNullOrEmpty(x.BinCutInstanceRow.PatternPinGroup));
            if (truthTables != null && truthTables.Any())
            {
                failFlags.AddRange(subFlowSheet.GetFailFlags());
                var harvRelatedFlags = new List<string>();
                if (TestPlanStatic.FlagOperationSheets.Any())
                {
                    foreach (FlagOperationSheet sheet in TestPlanStatic.FlagOperationSheets)
                    {
                        harvRelatedFlags.AddRange(sheet.GetAllFlags());
                    }
                }
                if (TestPlanStatic.AteStrSummarySheet != null)
                {
                    harvRelatedFlags.AddRange(TestPlanStatic.AteStrSummarySheet.GetAllFlags());
                }

                if (failFlags.Any(failFlag => harvRelatedFlags.Exists(x => x.Equals(failFlag, StringComparison.OrdinalIgnoreCase))))
                {
                    hasHarvItems = true;
                }
                HarvestingTruthTableSheet truthTable = truthTables.First();
                bool isHarvestFailFlagExist = false;
                List<Flag> flags = truthTable.GetDistinctFlags(failFlags);
                if (flags.Any() || hasHarvItems)
                {
                    isHarvestFailFlagExist = true;
                }

                if (isHarvestFailFlagExist)
                {
                    Regex.Replace(subFlowSheet.Name, "^Flow_", "", RegexOptions.IgnoreCase);
                    List<string> allFlags = truthTable.GetHarvAllFlagList();
                    if (TestPlanStatic.FlagOperationSheets.Any())
                    {
                        foreach (FlagOperationSheet sheet in TestPlanStatic.FlagOperationSheets)
                        {
                            foreach (string flag in sheet.GetAllFlags())
                            {
                                if (!allFlags.Exists(x => x.Equals(flag, StringComparison.OrdinalIgnoreCase)))
                                {
                                    allFlags.Add(flag);
                                }
                            }
                        }
                    }
                }
            }
        }

        protected virtual List<FlowRow> GenFlowBody(IGrouping<string, BinCutFinalInstanceRow> group)
        {
            var flowRows = new List<FlowRow>();
            foreach (BinCutFinalInstanceRow dataRow in group)
            {
                List<FlowRow> testFlowRows = new List<FlowRow>();
                testFlowRows.Add(GetTestRow(dataRow));
                if (!dataRow.BinCutInstanceRow.IsBypassBinOut)
                {
                    testFlowRows.Add(GetControlBinRow(dataRow));
                }
                List<FlowRow> ifFlowRows = new ScanNonBinCutInstance().GetIfFlowRows(dataRow, testFlowRows, null);
                flowRows.AddRange(ifFlowRows);
            }
            return flowRows;
        }

        public virtual string GetPayloadType(string pattern)
        {
            if (PayloadTypeTable == null || string.IsNullOrWhiteSpace(pattern))
            {
                return string.Empty;
            }

            foreach (DataRow row in PayloadTypeTable.Rows)
            {
                bool allColumnsMatch = true;
                for (int j = 1; j < PayloadTypeTable.Columns.Count; j++)
                {
                    string matchPattern = row[j].ToString();
                    string columnName = PayloadTypeTable.Columns[j].ColumnName;
                    string subName = GetSubName(pattern, columnName);
                    if (!subName.Equals(matchPattern, StringComparison.CurrentCultureIgnoreCase))
                    {
                        allColumnsMatch = false;
                        break;
                    }
                }

                if (allColumnsMatch)
                {
                    return row[0].ToString();
                }
            }

            return string.Empty;
        }

        internal static string GetSubName(string name, string rule)
        {
            if (string.Equals(rule, "full", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(name))
            {
                return name;
            }

            if (string.IsNullOrWhiteSpace(rule))
            {
                return string.Empty;
            }

            string[] words = name.Split('_');

            IEnumerable<string> selectedWords = rule.Split(',').Select(r => int.TryParse(r.Trim(), out int index) ? index : -1).Where(index => index >= 0 && index < words.Length).Select(index => words[index]);

            return string.Join("_", selectedWords);
        }

        internal virtual FlowRow GetTestRow(BinCutFinalInstanceRow row, string flowParameter = "")
        {
            string pattern = row.PayloadList == null || !row.PayloadList.Any() ? row.PatternList.First() : row.PayloadList.First();
            string payloadType = GetPayloadType(pattern);
            var flowRow = new FlowRow();
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                flowRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
            }
            flowRow.Opcode = row.NopByEnableWord ? OpCode.Nop : OpCode.Test;
            flowRow.Parameter = string.IsNullOrEmpty(flowParameter) ? row.GetParameter() : flowParameter;
            flowRow.Job = row.GetJob();
            flowRow.Enable = row.GetEnable();
            flowRow.Env = GetEnv(row);
            IEnumerable<string> failFlags = row.GetFlowFailAction(payloadType, row.PerformanceMode, flowRow.Parameter);
            flowRow.FailAction = string.Join(",", failFlags.Where(x => !x.StartsWith("!")));
            flowRow.PassAction = string.Join(",", failFlags.Where(x => x.StartsWith("!")).Select(x => x.TrimStart('!')));
            flowRow.IsSsn = row.IsSsn();
            return flowRow;
        }

        internal FlowRow GetControlBinRow(BinCutFinalInstanceRow row, string flowParameter = "")
        {
            HashSet<string> binTableRowsSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            binTableRowsSet.AddRange(GetBinOutDatas([row]).SelectMany(AtpgBinOutService.GetBinNameList));
            HashSet<string> controlBinFlags = binTableRowsSet.Select(x => $"Control{x}").ToHashSet(StringComparer.OrdinalIgnoreCase);
            var flowRow = new FlowRow();
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                flowRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
            }
            flowRow.Job = row.GetBinOutStage();
            flowRow.Enable = row.BinCutInstanceRow.BintableEnableWd;
            flowRow.Opcode = OpCode.FlagTrue;
            flowRow.Parameter = string.Join(",", controlBinFlags);
            flowRow.DeviceCondition = "Flag-true";
            flowRow.DeviceName = row.GetTestInstanceFailFlag(flowParameter);
            _controlBinOutAndInstanceFlags.Add(row.GetTestInstanceFailFlag(flowParameter));
            _controlBinOutAndInstanceFlags.AddRange(controlBinFlags);
            return flowRow;
        }

        internal string GetEnv(BinCutFinalInstanceRow binCutFinalInstanceRow)
        {
            var envs = new List<string>();
            if (binCutFinalInstanceRow.Nop)
            {
                return "BlankInstance";
            }

            if (binCutFinalInstanceRow.PatternList.Count != 0)
            {
                List<string> list = binCutFinalInstanceRow.PatternList;
                if (binCutFinalInstanceRow.PatternList.Exists(x => x.Contains("+")))
                {
                    list = new List<string>();
                    foreach (string pat in binCutFinalInstanceRow.PatternList)
                    {
                        if (pat.Contains("+"))
                        {
                            list.AddRange(pat.Split(new[] { '+', ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList());
                        }
                        else
                        {
                            list.Add(pat);
                        }
                    }
                }

                foreach (string pat in list)
                {
                    string temEnv = AcTSetCategoryMapSingleton.Instance().CheckAllPatternExist(pat);
                    if (!string.IsNullOrEmpty(temEnv))
                    {
                        envs.Add(temEnv);
                    }
                }
            }
            envs = envs.Distinct().ToList();
            return string.Join(",", envs);
        }

        internal List<PatSet> AddCommandAndFlagInPatSet(List<PatSet> patSetSheets, List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            var patSetHashSet = new HashSet<string>(binCutFinalInstanceRows.Where(x => x.PatSet != null && !string.IsNullOrEmpty(x.PatSet.PatSetName)).Select(x => x.PatSet.PatSetName), StringComparer.CurrentCultureIgnoreCase);
            var initPatSetHashSet = new HashSet<string>(binCutFinalInstanceRows.Where(x => x.InitPatSet != null && !string.IsNullOrEmpty(x.InitPatSet.PatSetName)).Select(x => x.InitPatSet.PatSetName), StringComparer.CurrentCultureIgnoreCase);
            foreach (string initPatSet in initPatSetHashSet)
            {
                patSetHashSet.Add(initPatSet);
            }

            for (int i = 0; i < patSetSheets.Count; i++)
            {
                if (!patSetHashSet.Contains(patSetSheets[i].PatSetName))
                {
                    patSetSheets[i].IsBackup = true;
                    foreach (PatSetRow row in patSetSheets[i].PatSetRows)
                    {
                        row.IsBackup = true;
                        row.Comment += ", dont_useInFlow";
                    }
                }

                foreach (PatSetRow item in patSetSheets[i].PatSetRows)
                {
                    Dictionary<string, PatternData> dictionary = AcTSetCategoryMapSingleton.Instance().PatternList;
                    if (string.IsNullOrEmpty(item.File))
                    {
                        continue;
                    }

                    string patternName = item.File.ToLower();
                    if (dictionary.ContainsKey(patternName))
                    {
                        if (dictionary[patternName].Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
                        {
                            item.Comment += ", dont_useInCsv";
                            item.IsBackup = true;
                        }
                        else
                        {
                            if (dictionary[patternName].FileVersion.Equals("n/a", StringComparison.OrdinalIgnoreCase))
                            {
                                item.Comment += ", no_pattern";
                                item.IsBackup = true;
                            }
                            else
                            {
                                if (!dictionary[patternName].IsExist)
                                {
                                    item.Comment += ", no_pattern";
                                    item.IsBackup = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        item.Comment += ", no_pattern";
                        item.IsBackup = true;
                    }
                }
            }
            return patSetSheets;
        }

        public List<string> GetAllFlags(string patternPinGroup)
        {
            string[] arr = patternPinGroup.Split(';');
            var flags = new List<string>();
            foreach (string item in arr)
            {
                if (_regex3.IsMatch(item))
                {
                    string flag = _regex3.Match(item).Groups["Flag"].ToString();
                    flags.Add(flag);
                }
            }
            return flags;
        }
    }
}
