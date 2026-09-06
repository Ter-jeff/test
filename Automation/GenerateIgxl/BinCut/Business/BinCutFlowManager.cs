using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlSheets;

using OfficeOpenXml;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.FlowNew;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutWorkFlowManager
    {
        private readonly BinCutInputData _binCutInputManager;

        //Output
        public BinCutPatternReport BinCutPatternReport;
        public BinCutPatternReport BinCutPostPatternReport;

        private List<BinCutFinalInstanceRow> _binCutFinalInstanceRows = new List<BinCutFinalInstanceRow>();
        private Dictionary<string, List<BinCutSourceItem>> _sourceDic = new Dictionary<string, List<BinCutSourceItem>>();

        public const string FlowVddBinningHvcc = "Flow_VddBinning_HVCC";
        private const string RegFlow = "^Flow_";
        private const string RegRelay = "^Relay_";
        public List<BinCutBinningItem> ExistHvccFlags;
        private HashSet<string> _binCutRegularPatSetRows = new HashSet<string>();

        public BinCutWorkFlowManager(BinCutInputData binCutInputManager)
        {
            _binCutInputManager = binCutInputManager;
        }

        public void Main()
        {
            GenBinCutFlow();

            if (_binCutInputManager.BinCutFlowTables.Any())
            {
                GenBinTable(_sourceDic, _binCutInputManager);
            }

            if (_binCutInputManager.BinCutPostFlowSheets.Any())
            {
                var binCutPostBinTableWriter = new BinCutPostBinTableWriter();
                binCutPostBinTableWriter.GetOutsideBinTable(_sourceDic, _binCutInputManager);
            }
        }

        public void GenBinCutFlow()
        {
            var binCutPreProcess = new BinCutPreProcess(_binCutInputManager);
            _binCutFinalInstanceRows = new List<BinCutFinalInstanceRow>();
            _binCutRegularPatSetRows = new HashSet<string>();
            List<string> gradeSearchJobs = GetAllGradeSearchJobs();

            #region bincut file
            if (_binCutInputManager.BinCutFlowTables.Any())
            {
                _binCutFinalInstanceRows = binCutPreProcess.GetBinCutInstanceRows(_binCutInputManager.BinCutInstanceSheets);
                List<BinCutSourceItem> sources = _binCutInputManager.NewBinCutFlowTables != null && _binCutInputManager.NewBinCutFlowTables.Any() ?
                    NewGetMergerSourceItem(_binCutInputManager.BinCutFlowTables, _binCutInputManager.NewBinCutFlowTables) :
                    GetMergerSourceItem(_binCutInputManager.BinCutFlowTables);

                _sourceDic = SplitByFlow(sources);

                _sourceDic = RemoveDoNotTest(_sourceDic);

                _sourceDic = RemoveHipSplitIsNull(_sourceDic, _binCutInputManager.NewBinCutFlowTables, _binCutInputManager.BinCutInstanceSheets);

                GetAcSpecSheet(_binCutFinalInstanceRows);

                BinCutFlowInstanceWriter binCutFlowInstanceWriter = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) ? new BinCutFlowInstanceWriterCs(_binCutInputManager, _binCutFinalInstanceRows, gradeSearchJobs) : new BinCutFlowInstanceWriter(_binCutInputManager, _binCutFinalInstanceRows, gradeSearchJobs);
                binCutFlowInstanceWriter.GetBinCutResult(_sourceDic, out BinCutPatternReport, out ExistHvccFlags);

                GetPatSetSheet(_binCutFinalInstanceRows);

                GetRegularPatSetRows(_binCutFinalInstanceRows, out _binCutRegularPatSetRows);
            }
            #endregion

            #region post bincut
            if (_binCutInputManager.BinCutPostFlowSheets.Any())
            {
                BinCutFinalInstanceRows binCutInstancePostRows = binCutPreProcess.GetBinCutInstanceRows(_binCutInputManager.BinCutInstancePostSheets, _binCutFinalInstanceRows);
                if (_binCutFinalInstanceRows != null && binCutInstancePostRows.Any())
                {
                    binCutInstancePostRows = binCutInstancePostRows.RePatSetNameDuplicateRowsByKey(_binCutFinalInstanceRows, "Post");
                }

                var sourceRowPostListDic = new Dictionary<string, List<BinCutSourceItem>>();
                foreach (BinCutFlowTables postFlowSheets in _binCutInputManager.BinCutPostFlowSheets)
                {
                    List<BinCutSourceItem> sourceRowPostList = _binCutInputManager.NewPostBinCutFlowTables != null && _binCutInputManager.NewPostBinCutFlowTables.Any() ?
                                                    NewGetMergerSourceItem(postFlowSheets, _binCutInputManager.NewPostBinCutFlowTables) :
                                                    GetMergerSourceItem(postFlowSheets);
                    sourceRowPostListDic.Add(postFlowSheets.First().SheetName, sourceRowPostList);
                }
                GetAcSpecSheet(binCutInstancePostRows);
                var postBinCutFlowInstanceWriter = new BinCutPostFlowInstanceWriter(_binCutInputManager, binCutInstancePostRows);
                postBinCutFlowInstanceWriter.GetBinCutOutsideResult(sourceRowPostListDic, out BinCutPostPatternReport);
                GetPatSetSheet(binCutInstancePostRows, _binCutRegularPatSetRows);
            }
            #endregion
        }

        internal Dictionary<string, List<BinCutSourceItem>> RemoveDoNotTest(Dictionary<string, List<BinCutSourceItem>> sourceDic)
        {
            const string interpolateDoNoTest = "_DONOTEST";
            const string interpolateDoNotTest = "_DONOTTEST";
            var newDic = new Dictionary<string, List<BinCutSourceItem>>();
            foreach (KeyValuePair<string, List<BinCutSourceItem>> item in sourceDic)
            {
                var list = item.Value.Where(x =>
                !(x.ColumnContent.EndsWith(interpolateDoNoTest, StringComparison.CurrentCultureIgnoreCase) ||
                x.ColumnContent.EndsWith(interpolateDoNotTest, StringComparison.CurrentCultureIgnoreCase))).ToList();
                newDic.Add(item.Key, list);
            }
            return newDic;
        }

        private Dictionary<string, List<BinCutSourceItem>> RemoveHipSplitIsNull(Dictionary<string, List<BinCutSourceItem>> sourceDic, NewBinCutFlowTables newBinCutFlowTables, List<BinCutInstanceSheet> binCutInstanceSheets)
        {
            var newHipResult = newBinCutFlowTables
                .Where(table => table.Rows != null)
                .SelectMany(table => table.Rows
                    .Where(row => row.SubFlowsByType != null)
                    .SelectMany(row => row.SubFlowsByType
                        .Where(subFlow => subFlow.Item2)
                        .Select(subFlow =>
                        {
                            string[] flowParts = subFlow.Item1?.Split(':');
                            string flowName = flowParts != null && flowParts.Length > 1 ? flowParts[1].Trim() : string.Empty;
                            return (JobName: table.JobName?.Trim() ?? string.Empty, FlowName: flowName);
                        })
                    )
                )
                .ToList();
            // If no newHIPResult, return sourceDic
            if (!newHipResult.Any())
            {
                return sourceDic;
            }
            // Get TTR return sourceDic
            var ttrResult = binCutInstanceSheets
                .Where(table => table.SheetName == "TTRTable")
                .SelectMany(table => table.Rows)
                .ToList();
            // Filter newHIPResult which not exists in TTR results
            newHipResult = newHipResult
                .Where(hip => !ttrResult.Any(ttr =>
                    string.Equals(ttr.FlowName.Trim(), hip.FlowName.Trim(), StringComparison.OrdinalIgnoreCase)
                ))
                .ToList();
            // Filter non-uses subflows 
            var newDic = new Dictionary<string, List<BinCutSourceItem>>();
            foreach (KeyValuePair<string, List<BinCutSourceItem>> item in sourceDic)
            {
                var list = item.Value.Where(x =>
                        !newHipResult.Any(hip =>
                            x.Job.Equals(hip.JobName.Trim(), StringComparison.CurrentCultureIgnoreCase) &&
                            x.ColumnContent.EndsWith(hip.FlowName.Trim(), StringComparison.CurrentCultureIgnoreCase)
                        )
                    ).ToList();
                newDic.Add(item.Key, list);
            }
            return newDic;
        }

        private void GenBinTable(Dictionary<string, List<BinCutSourceItem>> sourceRowDic, BinCutInputData binCutInputManager)
        {
            var binCutBinTableWriter = new BinCutBinTableWriter();
            binCutBinTableWriter.GenBinTable(sourceRowDic, binCutInputManager, ExistHvccFlags);
        }

        public void GetRegularPatSetRows(List<BinCutFinalInstanceRow> binCutFinalInstanceRows, out HashSet<string> regularPatSetRows)
        {
            if (binCutFinalInstanceRows == null)
            {
                regularPatSetRows = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                return;
            }

            var prevInitNames = new HashSet<string>(
                binCutFinalInstanceRows
                    .Where(p => !string.IsNullOrWhiteSpace(p.InitPatSetName))
                    .Select(p => p.InitPatSetName.Trim()),
                StringComparer.OrdinalIgnoreCase);

            var prevPatNames = new HashSet<string>(
                binCutFinalInstanceRows
                    .Where(p => !string.IsNullOrWhiteSpace(p.PatSetName))
                    .Select(p => p.PatSetName.Trim()),
                StringComparer.OrdinalIgnoreCase);

            regularPatSetRows = new HashSet<string>(prevInitNames, StringComparer.OrdinalIgnoreCase);
            regularPatSetRows.UnionWith(prevPatNames);
        }


        private void GetAcSpecSheet(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            AcSpecSheet acSpecSheet = TestProgram.IgxlWorkBk.GetAcSpecsSheet();
            if (acSpecSheet != null)
            {
                new BinCutAcSpecsWriter().GenAcSpecs(binCutFinalInstanceRows, acSpecSheet);
            }
        }

        private void GetPatSetSheet(List<BinCutFinalInstanceRow> binCutFinalInstanceRows, HashSet<string> regularPatSetRows = null)
        {
            if (!binCutFinalInstanceRows.Any())
            {
                return;
            }

            List<PatSetSheet> patSetSheets = new BinCutPatSetWriter().GetPatSetSheet(binCutFinalInstanceRows, regularPatSetRows);

            foreach (PatSetSheet patSetSheet in patSetSheets)
            {
                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirBinCut, patSetSheet);
            }

        }

        private List<string> GetAllGradeSearchJobs()
        {
            var allFlowTables = new BinCutFlowTables();
            if (_binCutInputManager.BinCutFlowTables.Any())
            {
                allFlowTables.AddRange(_binCutInputManager.BinCutFlowTables);
            }

            if (_binCutInputManager.BinCutFlowShadowTables.Any())
            {
                allFlowTables.AddRange(_binCutInputManager.BinCutFlowShadowTables);
            }

            if (_binCutInputManager.BinCutPostFlowSheets.Any())
            {
                allFlowTables.AddRange(_binCutInputManager.BinCutPostFlowSheets.SelectMany(x => x));
            }

            List<string> gradeSearchJobs = null;
            foreach (BinCutFlowTable binCutFlowTable in allFlowTables)
            {
                if (!binCutFlowTable.Rows.Exists(x => x.PinInfos.Any(pin => pin.PinContext.ContainsIgnoreCase("evaluate bin"))))
                {
                    continue;
                }

                if (gradeSearchJobs == null)
                {
                    gradeSearchJobs = new List<string>();
                }

                foreach (string job in binCutFlowTable.FinalJob.Where(job => !gradeSearchJobs.Contains(job)))
                {
                    gradeSearchJobs.Add(job);
                }
            }
            return gradeSearchJobs;
        }

        #region Source Row
        private List<BinCutSourceItem> NewGetMergerSourceItem(List<BinCutFlowTable> binCutSheets, NewBinCutFlowTables newBinCutFlowTables)
        {
            List<BinCutSourceItem> resultList = new List<BinCutSourceItem>();
            int jobCount = newBinCutFlowTables.JobList.Count;

            binCutSheets = binCutSheets.Where(x => !x.FinalJob.First().Equals("QA", StringComparison.CurrentCultureIgnoreCase)).ToList();

            if (!binCutSheets.Any(x => x.SheetName.ContainsIgnoreCase("post")))
            {
                if (IsHvFlowSplit(binCutSheets))
                {
                    MergeHvFlow(binCutSheets);
                }
            }

            foreach (IGrouping<string, NewBinCutFlowTable> sheetName in newBinCutFlowTables.GroupBy(x => x.SheetName))
            {
                NewBinCutFlowTables newBinCutFlowTable = new NewBinCutFlowTables();
                newBinCutFlowTable.AddRange(newBinCutFlowTables.FindAll(x => x.SheetName.Equals(sheetName.Key)));
                List<BinCutSourceItem> resultTmp = GetSourcesByRow(newBinCutFlowTable, binCutSheets);
                resultTmp = resultTmp.Select(p => { p.JobCount = jobCount; return p; }).ToList();
                resultList.AddRange(resultTmp);
            }
            return resultList;
        }

        private bool TryAddBinCutSource(
            List<BinCutSourceItem> binCutSourceItemsByRow,
            List<string> subflowTokens,
            string[] targetList,
            string subFlow,
            EnumColumnName colName,
            NewBinCutFlowTable newBinCutFlowTable,
            BinCutFlowSheetRow flowRow,
            NewBinCutFlowSheetRow row
        )
        {
            foreach (string item in targetList)
            {
                if (subflowTokens.Exists(x => x.Equals(item, StringComparison.CurrentCultureIgnoreCase)))
                {
                    binCutSourceItemsByRow.Add(
                        InitialSourceRowsNew(row, subFlow, colName, newBinCutFlowTable.JobName, flowRow)
                    );
                    return true;
                }
            }
            return false;
        }

        private static readonly string[] _scanList = new string[] { "TD", "SSSBBIST", "SSBBIST", "SBST", "CXT", "CDD", "CUX" };
        private static readonly string[] _bistList = new string[] { "BIST", "MBIST", "SRT", "SFT", "WUS", "SP3", "24N" };

        /// <summary>
        /// This rule array is order-sensitive. So please do not change order unless the check rule 
        /// has changed.
        /// </summary>
        private static readonly (
            Func<List<string>, bool> Match,
            EnumColumnName TargetColumn
        )[] _fallbackRules =
        {
            (arr => arr.Exists(x => x.Equals("ELB", StringComparison.CurrentCultureIgnoreCase)), EnumColumnName.ELB),
            (arr => arr.Exists(x => x.Equals("ILB", StringComparison.CurrentCulture)), EnumColumnName.ILB),
            (arr => arr.Exists(x => x.Equals("Set_E1_Voltage")), EnumColumnName.E1Voltage),
            (arr => arr.Exists(x => Regex.IsMatch(x, RegFlow) && x.ContainsIgnoreCase("ENABLE")), EnumColumnName.CallNwireEnable),
            (arr => arr.Exists(x => Regex.IsMatch(x, RegFlow) && x.ContainsIgnoreCase("DISABLE")), EnumColumnName.CallNwireDisable),
            (arr => arr.Exists(x => Regex.IsMatch(x, RegRelay) && x.ContainsIgnoreCase("ON")), EnumColumnName.RelayOn),
            (arr => arr.Exists(x => Regex.IsMatch(x, RegRelay) && x.ContainsIgnoreCase("OFF")), EnumColumnName.RelayOff),
            (arr => arr.Exists(x => Regex.IsMatch(x, RegFlow) && x.ContainsIgnoreCase("TMPS")), EnumColumnName.CallTMPS),
        };

        internal static EnumColumnName ResolveFallbackColumn(List<string> arr)
        {
            foreach ((Func<List<string>, bool> match, EnumColumnName column) in _fallbackRules)
            {
                if (match(arr))
                {
                    return column;
                }
            }

            return EnumColumnName.FUNC;
        }

        private List<BinCutSourceItem> GetBinCutSourceItems(
            NewBinCutFlowSheetRow row,
            List<BinCutFlowTable> binCutSheets,
            NewBinCutFlowTable newBinCutFlowTable)
        {
            var binCutSourceItems = new List<BinCutSourceItem>();
            foreach (string subFlow in row.SubFlows)
            {
                List<string> arr = subFlow.Split(new[] { ' ', '#', ':' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                string additionalMode = arr.First();
                if (!additionalMode.Split('_').First().Equals(row.PerformanceMode, StringComparison.CurrentCultureIgnoreCase))
                {
                    additionalMode = row.PerformanceMode;
                    ErrorReportManager.AddError(BinCutErrorType.E_FormatError_01, row.SheetName, row.RowNum, 0, $"The subflow {subFlow} had not start with Performance Mode !!!", new string[] { subFlow });
                }

                BinCutFlowSheetRow flowRow = binCutSheets
                    .SelectMany(x => x.Rows)
                    .ToList()
                    .Find(x =>
                        x.Job.Contains(row.Job)
                        && x.TableType.Equals(row.TableType)
                        && x.TableBinType.Equals(row.TableBinType)
                        && x.PerformanceMode.Equals(
                            additionalMode,
                            StringComparison.CurrentCultureIgnoreCase
                        )
                    );

                if (flowRow == null)
                {
                    string errorMessage =
                    $"Can not find voltage condition in flow sheet (Job: {row.Job}, TableType: {row.TableType} ,TableBinType: {row.TableBinType}, mode: {additionalMode})!!!";
                    ErrorReportManager.AddError(BinCutErrorType.E_FormatError_02, row.SheetName, row.RowNum, 0, $"Can not find voltage condition in flow sheet (Job: {row.Job}, TableType: {row.TableType} ,TableBinType: {row.TableBinType}, mode: {additionalMode})!!!", new string[] { row.Job, row.TableType.ToString(), row.TableBinType.ToString(), additionalMode });
                    continue;
                }
                bool found =
                    TryAddBinCutSource(binCutSourceItems, arr, _scanList, subFlow, EnumColumnName.TD, newBinCutFlowTable, flowRow, row)
                    || TryAddBinCutSource(binCutSourceItems, arr, _bistList, subFlow, EnumColumnName.Mbist, newBinCutFlowTable, flowRow, row);

                if (!found)
                {
                    EnumColumnName finalColName = ResolveFallbackColumn(arr);
                    binCutSourceItems.Add(InitialSourceRowsNew(row, subFlow, finalColName, newBinCutFlowTable.JobName, flowRow));
                }
            }
            return binCutSourceItems;
        }

        private List<BinCutSourceItem> GetSourcesByRow(NewBinCutFlowTables newBinCutFlowTables, List<BinCutFlowTable> binCutSheets)
        {
            var binCutSourceItems = new List<BinCutSourceItem>();
            IOrderedEnumerable<int> ints = newBinCutFlowTables
                .SelectMany(x => x.Rows)
                .Select(x => x.RowNum)
                .Distinct()
                .ToList()
                .OrderBy(x => x);
            Dictionary<string, int> subFlowMappingTable = newBinCutFlowTables
                .Find(x => x.SubFlowMappingTable.Any())
                .SubFlowMappingTable;
            foreach (int rowNum in ints)
            {
                var binCutSourceItemsByRow = new List<BinCutSourceItem>();
                foreach (NewBinCutFlowTable newBinCutFlowTable in newBinCutFlowTables)
                {
                    NewBinCutFlowSheetRow row = newBinCutFlowTable.Rows.Find(x => x.RowNum == rowNum);
                    if (row == null)
                    {
                        continue;
                    }

                    binCutSourceItemsByRow.AddRange(
                        GetBinCutSourceItems(row, binCutSheets, newBinCutFlowTable)
                    );
                }
                var byRow = binCutSourceItemsByRow
                    .OrderBy(x => subFlowMappingTable.TryGetValue(x.ColumnContent, out int value) ? value : -1)
                    .ToList();
                SetCanBeMerged(byRow, subFlowMappingTable);
                binCutSourceItems.AddRange(byRow);
            }
            return binCutSourceItems;
        }

        internal List<BinCutSourceItem> GetMergerSourceItem(List<BinCutFlowTable> binCutSheets)
        {

            binCutSheets = binCutSheets.Where(x => !x.FinalJob.First().Equals("QA", StringComparison.CurrentCultureIgnoreCase)).ToList();

            SortRowsByBinCutOrder(binCutSheets);

            if (!binCutSheets.Any(x => x.SheetName.ContainsIgnoreCase("post")))
            {
                if (IsHvFlowSplit(binCutSheets))
                {
                    MergeHvFlow(binCutSheets);
                }
            }

            int jobNumber = binCutSheets.Count;
            var resultList = new List<BinCutSourceItem>();
            var ints = binCutSheets.SelectMany(x => x.Rows).Select(x => x.RowNum).Distinct().ToList();
            // Set CanBeMerged
            foreach (int rowNum in ints)
            {
                var rows = binCutSheets.SelectMany(y => y.Rows.Where(x => x.RowNum.Equals(rowNum))).ToList();
                int atpg = rows.Select(x => x.Atpg).Where(x => !string.IsNullOrEmpty(x)).Distinct().Count();
                int mbist = rows.Select(x => x.Mbist).Where(x => !string.IsNullOrEmpty(x)).Distinct().Count();
                int spiRtos = rows.Select(x => x.SpiRtos).Where(x => !string.IsNullOrEmpty(x)).Distinct().Count();
                foreach (BinCutFlowSheetRow row in rows)
                {
                    if (atpg == 1)
                    {
                        row.CanBeMergedAtpg = true;
                    }

                    if (mbist == 1)
                    {
                        row.CanBeMergedMbist = true;
                    }

                    if (spiRtos == 1)
                    {
                        row.CanBeMergedSpiRtos = true;
                    }
                }
            }

            foreach (int rowNum in ints)
            {
                resultList.AddRange(GetSourcesByRow(binCutSheets, rowNum));
            }

            //Set all JobNumber as max job count
            resultList = resultList.Select(p => { p.JobCount = jobNumber; return p; }).ToList();
            return resultList;
        }

        private List<BinCutSourceItem> GetSourcesByRow(List<BinCutFlowTable> binCutSheets, int rowNum)
        {
            var oneRowNew = new List<BinCutSourceItem>();
            var tdSources = new List<BinCutSourceItem>();
            var mbistSources = new List<BinCutSourceItem>();
            var funcSources = new List<BinCutSourceItem>();
            string performanceMode = "";
            foreach (BinCutFlowTable binCutSheet in binCutSheets)
            {
                BinCutFlowSheetRow row = binCutSheet.Rows.Find(x => x.RowNum.Equals(rowNum));
                foreach (string finalJob in binCutSheet.FinalJob)
                {
                    if (row != null)
                    {
                        performanceMode = row.PerformanceMode;
                        tdSources.AddRange(InitialSourceRows(row, row.Atpg, EnumColumnName.TD, finalJob, row.PinInfos, row.CanBeMergedAtpg));
                        mbistSources.AddRange(InitialSourceRows(row, row.Mbist, EnumColumnName.Mbist, finalJob, row.PinInfos, row.CanBeMergedMbist));
                        funcSources.AddRange(InitialSourceRows(row, row.SpiRtos, EnumColumnName.FUNC, finalJob, row.PinInfos, row.CanBeMergedSpiRtos));
                    }
                }
            }

            #region sort row
            if (_binCutInputManager.BinCutOrderSheet != null)
            {
                if (_binCutInputManager.BinCutOrderSheet.Rows.Exists(x => x.PerformanceMode.Equals(performanceMode, StringComparison.CurrentCultureIgnoreCase)))
                {
                    BinCutOrderRow binCutOrderRow = _binCutInputManager.BinCutOrderSheet.Rows.Find(x => x.PerformanceMode.Equals(performanceMode, StringComparison.CurrentCultureIgnoreCase));
                    if (binCutOrderRow.Td != null && binCutOrderRow.Bist != null && binCutOrderRow.Func != null)
                    {
                        var dic = new Dictionary<string, int?>
                        {
                            { "TD", binCutOrderRow.Td }, { "Bist", binCutOrderRow.Bist }, { "Func", binCutOrderRow.Func }
                        };
                        dic = dic.OrderBy(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
                        foreach (KeyValuePair<string, int?> instanceOrder in dic)
                        {
                            if (Equals(instanceOrder.Key, "TD"))
                            {
                                oneRowNew.AddRange(tdSources);
                            }
                            else if (Equals(instanceOrder.Key, "Bist"))
                            {
                                oneRowNew.AddRange(mbistSources);
                            }
                            else if (Equals(instanceOrder.Key, "Func"))
                            {
                                oneRowNew.AddRange(funcSources);
                            }
                        }
                        return oneRowNew;
                    }
                }
            }

            //Modify CanBeMerged
            //MSX001 SOC TD,MSX001 SOC TD,MSX001 CPU TD,MSX001 CPU TD => all CanBeMerged
            //MSX001 SOC TD,MSX001 CPU TD,MSX001 SOC TD => MSX001 CPU TD CanBeMerged
            //MSX001 SOC TD,MSX001 CPU TD,MSX001 SOC TD,MSX001 CPU TD => all cna not CanBeMerged
            SetCanBeMerged(tdSources);
            SetCanBeMerged(mbistSources);
            SetCanBeMerged(funcSources);

            oneRowNew.AddRange(tdSources);
            oneRowNew.AddRange(mbistSources);
            oneRowNew.AddRange(funcSources);
            #endregion

            return oneRowNew;
        }

        private static void SetCanBeMerged(List<BinCutSourceItem> tdSources, Dictionary<string, int> subFlowMappingTable = null)
        {
            var groups = tdSources.GroupBy(x => x.Job).ToList();
            var originalKeys = groups.SelectMany(x => x).Select(x => x.ColumnContent).Distinct().ToList();
            List<string> searchKeys = originalKeys;
            if (subFlowMappingTable != null)
            {
                searchKeys = originalKeys.OrderBy(x => subFlowMappingTable[x]).ToList();
            }
            foreach (IGrouping<string, BinCutSourceItem> group in groups)
            {
                int currentIndex = -1;
                foreach (BinCutSourceItem item in group)
                {

                    int index = searchKeys.FindIndex(x => item.ColumnContent.Equals(x, StringComparison.CurrentCultureIgnoreCase));
                    if (index > currentIndex)
                    {
                        item.CanBeMerged = true;
                        currentIndex = index;
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }

        private void SortRowsByBinCutOrder(List<BinCutFlowTable> binCutFlowSheets)
        {
            ExcelWorksheet sheet = EpWorkbook.TestPlanWorkbook.Worksheets["BinCutOrder"];
            if (sheet == null && EpWorkbook.BinCutWorkbook != null)
            {
                sheet = EpWorkbook.BinCutWorkbook.Worksheets["BinCutOrder"];
            }

            if (sheet != null)
            {
                var binCutOrderReader = new BinCutOrderReader(new BinCutFlowTables(binCutFlowSheets));
                _binCutInputManager.BinCutOrderSheet = binCutOrderReader.ReadSheet(sheet);
                foreach (BinCutFlowTable binCutSheet in binCutFlowSheets)
                {
                    binCutSheet.SortByBinCutOrder(_binCutInputManager.BinCutOrderSheet.Rows);
                }
            }
        }

        private bool IsHvFlowSplit(List<BinCutFlowTable> binCutSheets)
        {
            foreach (BinCutFlowTable sheet in binCutSheets)
            {
                var hvRows = sheet.Rows.Where(x => x.TableType == EnumBinCutTableType.Hv).ToList();

                foreach (BinCutFlowSheetRow row in hvRows)
                {
                    row.AllOther = Regex.Replace(row.AllOther, "^0$", "");
                    row.Atpg = Regex.Replace(row.Atpg, "^0$", "");
                    row.Mbist = Regex.Replace(row.Mbist, "^0$", "");
                    row.PerformanceMode = Regex.Replace(row.PerformanceMode, "^0$", "");
                    row.SpiRtos = Regex.Replace(row.SpiRtos, "^0$", "");
                }

                var groups = hvRows.GroupBy(x => x.TableBinType).ToList();
                if (groups.Count > 1)
                {
                    var flow = groups.First().ToList();
                    foreach (IGrouping<EnumBinCutTableBinType, BinCutFlowSheetRow> group in groups)
                    {
                        var newFlow = group.ToList();
                        if (newFlow.Count != flow.Count)
                        {
                            return false;
                        }

                        for (int i = 0; i < flow.Count; i++)
                        {
                            if (!flow[i].PerformanceMode.Equals(newFlow[i].PerformanceMode, StringComparison.CurrentCultureIgnoreCase))
                            {
                                return false;
                            }

                            if (!flow[i].AllOther.Equals(newFlow[i].AllOther, StringComparison.CurrentCultureIgnoreCase))
                            {
                                return false;
                            }

                            if (!flow[i].Atpg.Equals(newFlow[i].Atpg, StringComparison.CurrentCultureIgnoreCase))
                            {
                                return false;
                            }

                            if (!flow[i].Mbist.Equals(newFlow[i].Mbist, StringComparison.CurrentCultureIgnoreCase))
                            {
                                return false;
                            }

                            if (!flow[i].SpiRtos.Equals(newFlow[i].SpiRtos, StringComparison.CurrentCultureIgnoreCase))
                            {
                                return false;
                            }
                        }
                    }
                }
            }
            return true;
        }

        private void MergeHvFlow(List<BinCutFlowTable> binCutSheets)
        {
            foreach (BinCutFlowTable sheet in binCutSheets)
            {
                var rows = new List<BinCutFlowSheetRow>();
                var lvRows = sheet.Rows.Where(x => x.TableType == EnumBinCutTableType.Lv).ToList();
                var hvRows = sheet.Rows.Where(x => x.TableType == EnumBinCutTableType.Hv).ToList();
                var groups = hvRows.GroupBy(x => x.TableBinType).ToList();
                rows.AddRange(lvRows);
                if (groups.Any())
                {
                    rows.AddRange(groups.First());
                }

                sheet.Rows = rows;
            }
        }

        private BinCutSourceItem InitialSourceRowsNew(NewBinCutFlowSheetRow newBinCutFlowSheetRow, string subFlow, EnumColumnName columnName, string job, BinCutFlowSheetRow binCutFlowSheetRow)
        {
            var sourceRow = new BinCutSourceItem(binCutFlowSheetRow, newBinCutFlowSheetRow, columnName, subFlow)
            {
                CanBeMerged = false,
                Job = job
            };
            if (columnName == EnumColumnName.E1Voltage)
            {
                sourceRow.Enable = newBinCutFlowSheetRow.Enable;
            }
            return sourceRow;
        }

        private List<BinCutSourceItem> InitialSourceRows(BinCutFlowSheetRow sheetRow, string columnContent, EnumColumnName columnName,
            string job, List<PinInfo> pinInfos, bool canBeMerged)
        {
            //Ex-1: MC607 TD, MC407 TD
            //Ex-2: MS221 BIST (SRVB_GR03,SRVB_GR04)
            //Ex-3: SOC TD (except CUO0)
            //The delimiter is comma and underscore

            #region replace delimiter comma to @ for Ex-2
            var sourceRows = new List<BinCutSourceItem>();
            string newColumnContent = columnContent;
            if (columnContent.Contains(","))
            {
                const string regexPattern = "[(](?<str>.*)[)]";
                if (Regex.IsMatch(columnContent, regexPattern, RegexOptions.IgnoreCase))
                {
                    foreach (object item in Regex.Matches(columnContent, regexPattern, RegexOptions.IgnoreCase))
                    {
                        string oldName = item.ToString();
                        if (item.ToString().Contains(','))
                        {
                            string newName = Regex.Replace(oldName, ",", "@");
                            newColumnContent = columnContent.Replace(oldName, newName);
                        }
                    }
                }
            }
            #endregion
            List<string> contentList = Regex.Split(newColumnContent, ",|;").ToList();
            foreach (string content in contentList)
            {
                if (content.Trim() == "" || content == "0" || content == "0.0")
                {
                    continue;
                }

                string newContent = content.Replace("@", ",").Trim();
                var sourceRow = new BinCutSourceItem(sheetRow, columnName, newContent, pinInfos)
                {
                    CanBeMerged = canBeMerged,
                    Job = job
                };
                sourceRows.Add(sourceRow);
            }

            return sourceRows;
        }

        #region Serperate into different flows
        private Dictionary<string, List<BinCutSourceItem>> SplitByFlow(List<BinCutSourceItem> sourceRowList)
        {
            //SortBinCutSourceRowByTable(sourceRowList);
            List<string> extraFlow = _binCutInputManager.BinCutOrderSheet != null ? _binCutInputManager.BinCutOrderSheet.GetExtraFlow() : GetExtraMode();

            var sourceRowDic = new Dictionary<string, List<BinCutSourceItem>>();
            foreach (BinCutSourceItem row in sourceRowList)
            {
                string sheetName = GetSheetName(row, extraFlow);
                if (sourceRowDic.ContainsKey(sheetName))
                {
                    sourceRowDic[sheetName].Add(row);
                }
                else
                {
                    sourceRowDic.Add(sheetName, new List<BinCutSourceItem> { row });
                }
            }

            if (sourceRowDic.Count(x => x.Key.StartsWith(FlowVddBinningHvcc, StringComparison.CurrentCultureIgnoreCase)) == 1)
            {
                var newSourceRowDic = new Dictionary<string, List<BinCutSourceItem>>();
                foreach (KeyValuePair<string, List<BinCutSourceItem>> item in sourceRowDic)
                {
                    string name = item.Key.StartsWith(FlowVddBinningHvcc, StringComparison.CurrentCultureIgnoreCase) ? FlowVddBinningHvcc : item.Key;
                    newSourceRowDic.Add(name, item.Value);
                }
                return newSourceRowDic;
            }
            return sourceRowDic;
        }

        private List<string> GetExtraMode()
        {
            List<string> modeList;
            if (_binCutInputManager.NewBinCutFlowTables != null && _binCutInputManager.NewBinCutFlowTables.Any())
            {
                modeList = _binCutInputManager.NewBinCutFlowTables.First(x => !x.SheetName.Equals("HBV") && !x.SheetName.Equals("PBC")).Rows.Where(row => row.TableType == EnumBinCutTableType.Lv).Select(x => x.PerformanceMode).ToList();
            }
            else
            {
                modeList = _binCutInputManager.BinCutFlowTables.First().Rows.Where(row => row.TableType == EnumBinCutTableType.Lv).Select(x => x.PerformanceMode).ToList();
            }

            return BinCutExtraFlowName.GetExtraPerformanceMode(modeList);
        }

        private string GetSheetName(BinCutSourceItem row, List<string> extraFlow)
        {
            string sheetName = GetFlowSheetName(row);
            if (extraFlow.Exists(x => x.Equals(row.PerformanceMode, StringComparison.CurrentCultureIgnoreCase)))
            {
                sheetName = "Flow_" + row.PerformanceMode + "_TD_Mbist_BV";
            }

            return sheetName;
        }

        private string GetFlowSheetName(BinCutSourceItem row)
        {
            if (row.TableType == EnumBinCutTableType.Hv && row.TableBinType == EnumBinCutTableBinType.Bin1)
            {
                return FlowVddBinningHvcc;
            }

            if (row.TableType == EnumBinCutTableType.Hv && row.TableBinType == EnumBinCutTableBinType.BinX)
            {
                return FlowVddBinningHvcc + "_BinX";
            }

            if (row.TableType == EnumBinCutTableType.Hv && row.TableBinType == EnumBinCutTableBinType.BinY)
            {
                return FlowVddBinningHvcc + "_BinY";
            }

            return "Flow_" + row.TargetPerformanceMode + "_TD_Mbist_BV";
        }
        #endregion

        #endregion

    }
}
