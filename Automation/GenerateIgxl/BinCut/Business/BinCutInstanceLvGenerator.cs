using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business.BinCutInstance;
using Automation.GenerateIgxl.Scan.Harvest.Flow;
using Automation.InputManager.Data;
using Automation.Singleton;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutInstanceLvGenerator
    {
        protected IBinCutInstance InstanceInterface;

        public BinCutPatternSearch PatternSearch;
        public List<BinCutRow> InstanceJobsMapping = new List<BinCutRow>();
        protected List<BinCutBinningItem> ExistHvccFlags = new List<BinCutBinningItem>();
        protected List<BinCutBinningItem> ExistPostFlags = new List<BinCutBinningItem>();
        protected MultiTestSettingSheetsSingleton MultiTestSettingSheets;
        protected readonly BinCutInputData BinCutInputManager;
        protected string Type;

        public BinCutInstanceLvGenerator(List<BinCutSourceItem> sources, BinCutInputData binCutInputManager, List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            BinCutInputManager = binCutInputManager;
            foreach (BinCutSourceItem sourceRow in sources)
            {
                var data = new BinCutRow();
                PatternSearch = new BinCutPatternSearch(sourceRow, binCutFinalInstanceRows);
                data.BinCutFinalInstanceRows = PatternSearch.GetPatterns();
                data.BinCutSourceItem = sourceRow;
                InstanceJobsMapping.Add(data);
            }

            GenBlankRow(InstanceJobsMapping);
            MultiTestSettingSheets = MultiTestSettingSheetsSingleton.Instance();
            Type = "BV";
        }

        public List<BinCutBinningItem> GetHvccFlags()
        {
            return ExistHvccFlags;
        }

        public void GenInstanceRows(bool isPost, BinCutInputData binCutInputData)
        {
            foreach (BinCutRow instanceJob in InstanceJobsMapping)
            {
                var instanceRowDetail = new List<InstanceRow>();
                foreach (BinCutFinalInstanceRow dataRow in instanceJob.BinCutFinalInstanceRows)
                {
                    InstanceInterface = new BinCutInstanceInterfaceFactory().GetBinCutInstanceInterface(dataRow, instanceJob.BinCutSourceItem, binCutInputData, isPost);
                    List<InstanceRow> rows = GetInstanceRow(dataRow);
                    if (rows != null && rows.Count != 0)
                    {
                        instanceRowDetail.AddRange(rows);
                    }
                }
                instanceJob.InstanceRowDetail = instanceRowDetail;
            }
        }

        public void BinCutInstanceRowMergeByJob()
        {
            for (int i = 0; i < InstanceJobsMapping.Count; i++)
            {
                if (!InstanceJobsMapping[i].BinCutSourceItem.CanBeMerged)
                {
                    continue;
                }

                List<InstanceRow> row1 = InstanceJobsMapping[i].InstanceRowDetail;
                for (int j = i + 1; j < InstanceJobsMapping.Count; j++)
                {
                    if (!InstanceJobsMapping[j].BinCutSourceItem.CanBeMerged)
                    {
                        continue;
                    }

                    List<InstanceRow> row2 = InstanceJobsMapping[j].InstanceRowDetail;
                    for (int m = 0; m < row1.Count; m++)
                    {
                        for (int n = 0; n < row2.Count; n++)
                        {
                            if (!row1[m].GetDifferences(row2[n]).Any() && row1[m].RowNum.Equals(row2[n].RowNum))
                            {
                                if (InstanceJobsMapping[i].BinCutSourceItem.ColumnContent != InstanceJobsMapping[j].BinCutSourceItem.ColumnContent)
                                {
                                    break;
                                }

                                row2.Remove(row2[n]);
                                if (InstanceJobsMapping[i].BinCutFinalInstanceRows[m].JobNotMap != InstanceJobsMapping[j].BinCutFinalInstanceRows[n].JobNotMap)
                                {
                                    InstanceJobsMapping[i].BinCutFinalInstanceRows[m].JobNotMap = false;
                                }

                                InstanceJobsMapping[j].BinCutFinalInstanceRows.Remove(InstanceJobsMapping[j].BinCutFinalInstanceRows[n]);
                                if (!InstanceJobsMapping[i].BinCutFinalInstanceRows[m].FinalJobs.Exists(x => x.Equals(InstanceJobsMapping[j].BinCutSourceItem.Job)))
                                {
                                    InstanceJobsMapping[i].BinCutFinalInstanceRows[m].FinalJobs.Add(InstanceJobsMapping[j].BinCutSourceItem.Job);
                                }

                                if (row2.Count == 0)
                                {
                                    InstanceJobsMapping.Remove(InstanceJobsMapping[j]);
                                    j = i;
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }

        public void BinCutInstanceNameCheck()
        {
            var instanceRowDetails = InstanceJobsMapping.SelectMany(x => x.InstanceRowDetail.
                Select(y => new BinCutInstanceForReName { BinCutFinalInstanceRows = x.BinCutFinalInstanceRows, InstanceRowDetail = y })).ToList();

            bool existDuplicate = false;
            for (int i = 0; i < instanceRowDetails.Count; i++)
            {
                BinCutInstanceForReName row1 = instanceRowDetails[i];
                for (int j = i + 1; j < instanceRowDetails.Count; j++)
                {
                    BinCutInstanceForReName row2 = instanceRowDetails[j];
                    if (!row1.InstanceRowDetail.TestName.Equals(row2.InstanceRowDetail.TestName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        continue;
                    }

                    List<string> diffTypes = row1.InstanceRowDetail.GetDifferences(row2.InstanceRowDetail);
                    if (!diffTypes.Any())
                    {
                        continue;
                    }

                    List<string> initList = row2.InstanceRowDetail.InitList;
                    List<string> payloadList = row2.InstanceRowDetail.PayloadList;
                    foreach (BinCutFinalInstanceRow instance in row2.BinCutFinalInstanceRows)
                    {
                        if (row2.InstanceRowDetail.RowNum.Equals(instance.BinCutInstanceRow.RowNum) &&
                            initList.Count == instance.InitList.Count && payloadList.Count == instance.PayloadList.Count)
                        {
                            string initstr = string.Join("-", initList);
                            string initstr2 = string.Join("-", instance.InitList);
                            string payloadstr = string.Join("-", payloadList);
                            string payloadstr2 = string.Join("-", instance.PayloadList);
                            if (initstr.Equals(initstr2, StringComparison.CurrentCultureIgnoreCase) &&
                                payloadstr.Equals(payloadstr2, StringComparison.CurrentCultureIgnoreCase))
                            {
                                existDuplicate = true;
                                instance.IsDuplicateName = true;
                                instance.FinalInstName = Combination.CombineByUnderLine(instance.PatSetName, string.Join("_", diffTypes));
                            }
                        }
                    }
                }
            }

            if (existDuplicate)
            {
                var instanceRows = new BinCutFinalInstanceRows();
                instanceRows.AddRange(InstanceJobsMapping.SelectMany(x => x.BinCutFinalInstanceRows.Where(y => !string.IsNullOrEmpty(y.FinalInstName))));
                instanceRows.RePatSetNameDuplicateRowsBySerialNum();
            }
        }

        public List<BinCutRowForSort> ReArrangeByOrderOption(string sheetName, bool orderByInst)
        {
            var binCutRows = InstanceJobsMapping.SelectMany(x => x.BinCutFinalInstanceRows.
                Select(y => new BinCutRowForSort { BinCutSourceRow = x.BinCutSourceItem, BinCutFinalInstanceRow = y })).ToList();

            bool isOrderByPmOrder = GetIsOrderByPmOrder(binCutRows);
            if (isOrderByPmOrder)
            {
                var binCutRowsNew = new List<BinCutRowForSort>();
                var withBinCutInstDataRows = binCutRows.Where(x => x.BinCutFinalInstanceRow.BinCutInstanceRow != null).ToList();
                IEnumerable<BinCutRowForSort> withOutBinCutInstDataRows = binCutRows.Where(x => x.BinCutFinalInstanceRow.BinCutInstanceRow == null);
                List<string> existSheetNames = GetExistSheetNames();
                var withOutPmOrderRows = new List<BinCutRowForSort>();
                var withPmOrderRows = new List<BinCutRowForSort>();
                foreach (string existSheetName in existSheetNames)
                {
                    var rows = withBinCutInstDataRows.Where(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.SheetName.Equals(existSheetName)).ToList();
                    rows = rows.Where(x => !x.BinCutFinalInstanceRow.BinCutInstanceRow.RowNum.Equals(0)).OrderBy(y => y.BinCutFinalInstanceRow.BinCutInstanceRow.RowNum).ToList();

                    var withPmOrderRowsTmp = rows.Where(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.PmOrder != -1 && !string.IsNullOrEmpty(x.BinCutFinalInstanceRow.BinCutInstanceRow.PerfMode)).ToList();
                    // check if PerfMode equal mode of source row
                    if (!sheetName.StartsWith(BinCutWorkFlowManager.FlowVddBinningHvcc, StringComparison.OrdinalIgnoreCase))
                    {
                        withPmOrderRows.AddRange(withPmOrderRowsTmp.Where(x => x.BinCutSourceRow.PerformanceMode.Split('_').First().Equals(x.BinCutFinalInstanceRow.BinCutInstanceRow.PerfMode, StringComparison.CurrentCultureIgnoreCase)).ToList());
                    }
                    else
                    {
                        withPmOrderRows.AddRange(withPmOrderRowsTmp.Where(x => x.BinCutSourceRow.PerformanceMode.Split('_').First().Equals(x.BinCutFinalInstanceRow.BinCutInstanceRow.PerfMode.Split('_').First(), StringComparison.CurrentCultureIgnoreCase)).ToList());
                    }
                    withOutPmOrderRows.AddRange(rows.Where(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.PmOrder == -1 || string.IsNullOrEmpty(x.BinCutFinalInstanceRow.BinCutInstanceRow.PerfMode)).ToList());
                }
                IEnumerable<IGrouping<string, BinCutRowForSort>> groups = withPmOrderRows.GroupBy(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.PerfMode.Split('_').First());
                foreach (IGrouping<string, BinCutRowForSort> group in groups)
                {
                    binCutRowsNew.AddRange(group.OrderBy(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.PmOrder).ToList());
                }

                binCutRowsNew.AddRange(withOutPmOrderRows);
                binCutRowsNew.AddRange(withBinCutInstDataRows.Where(x => !existSheetNames.Any(y => y.Equals(x.BinCutFinalInstanceRow.BinCutInstanceRow.SheetName))).ToList());
                binCutRowsNew.AddRange(withOutBinCutInstDataRows);
                binCutRows = binCutRowsNew;
            }
            else if (orderByInst)
            {
                var binCutRowsNew = new List<BinCutRowForSort>();
                var chunks = binCutRows.ChunkBy(x => x.BinCutSourceRow.PerformanceMode.Split('_').First()).ToList();
                foreach (IGrouping<string, BinCutRowForSort> chunk in chunks)
                {
                    binCutRowsNew.AddRange(chunk.Where(x => !x.BinCutFinalInstanceRow.BinCutInstanceRow.RowNum.Equals(0)).OrderBy(y => y.BinCutFinalInstanceRow.BinCutInstanceRow.RowNum).ToList());
                    binCutRowsNew.AddRange(chunk.Where(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.RowNum.Equals(0)));
                }
                binCutRows = binCutRowsNew;
            }
            else
            {
                var binCutRowsNew = new List<BinCutRowForSort>();
                foreach (BinCutRowForSort binCutRow in binCutRows)
                {
                    //Need to sort by rowNum when merge jobs
                    if (binCutRow.BinCutSourceRow.CanBeMerged)
                    {
                        if (!binCutRowsNew.Exists(x =>
                            x.BinCutSourceRow.ColumnContent.Equals(binCutRow.BinCutSourceRow.ColumnContent, StringComparison.CurrentCultureIgnoreCase) &&
                            x.BinCutSourceRow.ColumnName == binCutRow.BinCutSourceRow.ColumnName))
                        {
                            var rows = binCutRows.Where(x =>
                                x.BinCutSourceRow.ColumnContent.Equals(binCutRow.BinCutSourceRow.ColumnContent, StringComparison.CurrentCultureIgnoreCase) &&
                                x.BinCutSourceRow.ColumnName == binCutRow.BinCutSourceRow.ColumnName && x.BinCutSourceRow.CanBeMerged).ToList();
                            var rowsWithRowNum = rows.Where(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.RowNum > 0).ToList();
                            rowsWithRowNum = rowsWithRowNum.OrderBy(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.RowNum).ToList();
                            var rowsWithoutRowNum = rows.Where(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.RowNum <= 0).ToList();
                            binCutRowsNew.AddRange(rowsWithRowNum);
                            binCutRowsNew.AddRange(rowsWithoutRowNum);
                        }
                    }
                    else
                    {
                        binCutRowsNew.Add(binCutRow);
                    }
                }
                binCutRows = binCutRowsNew;
            }
            return binCutRows;
        }

        public List<InstanceRow> GetInstanceRows(string flowName, List<BinCutRowForSort> binCutRows, out List<BinCutPatternRow> binCutPatternRows, bool isPost, bool intSkipTest)
        {
            binCutPatternRows = new List<BinCutPatternRow>();
            var instanceRowList = new List<InstanceRow>();
            for (int index = 0; index < binCutRows.Count; index++)
            {
                BinCutSourceItem sourceRow = binCutRows[index].BinCutSourceRow;
                BinCutFinalInstanceRow dataRow = binCutRows[index].BinCutFinalInstanceRow;
                if (sourceRow.ColumnContent.StartsWith("Flow_TMPS", StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                if (sourceRow.InstOrCallFlowByBms && !sourceRow.ColumnName.Equals(EnumColumnName.E1Voltage))
                {
                    continue;
                }

                InstanceInterface = new BinCutInstanceInterfaceFactory().GetBinCutInstanceInterface(dataRow, sourceRow, BinCutInputManager, isPost);
                List<InstanceRow> rows = GetInstanceRow(dataRow);
                if (rows != null && rows.Count != 0)
                {
                    instanceRowList.AddRange(rows);
                }

                foreach (InstanceRow row in rows)
                {
                    if (dataRow.BinCutInstanceRow.Type == BincutInstanceType.Hardip || dataRow.BinCutInstanceRow.Type == BincutInstanceType.Rtos)// (!string.IsNullOrEmpty(dataRow.TestName))
                    {
                        foreach (string pattern in dataRow.PatternList)
                        {
                            var binCutPatternRow = new BinCutPatternRow(flowName)
                            {
                                RowNum = dataRow.BinCutInstanceRow.RowNum,
                                BinCutInstanceRow = dataRow.BinCutInstanceRow,
                                Voltage = Type,
                                BinningDomain = sourceRow.BinningDomain,
                                PerformanceMode = sourceRow.PerformanceMode,
                                Type = sourceRow.ColumnName.ToString(),
                                Condition = sourceRow.ColumnContent,
                                DeviceCondition = GetDeviceCondition(dataRow),
                                InstanceName = row.TestName,
                                Jobs = row.FinalJobs,
                                Pattern = pattern,
                                Instance = dataRow.BinCutInstanceRow.Instance,
                                IsInterpolateSkip = intSkipTest
                            };
                            if (!string.IsNullOrEmpty(pattern.Trim()))
                            {
                                binCutPatternRow.SheetName = dataRow.BinCutInstanceRow.SheetName;
                                binCutPatternRows.Add(binCutPatternRow);
                            }
                        }
                    }
                    else if (sourceRow.ColumnName == EnumColumnName.Mbist)
                    {
                        foreach (string pattern in dataRow.InitList)
                        {
                            BinCutPatternRow binCutPatternRow = GetBinCutPatternRow(flowName, dataRow, row, sourceRow, pattern, intSkipTest);
                            binCutPatternRow.PattrenSetName = dataRow.InitPatSetName ?? "";
                            binCutPatternRow.SheetName = dataRow.BinCutInstanceRow.SheetName;
                            if (!string.IsNullOrEmpty(pattern.Trim()))
                            {
                                binCutPatternRows.Add(binCutPatternRow);
                            }
                        }
                        foreach (string pattern in dataRow.PayloadList)
                        {
                            BinCutPatternRow binCutPatternRow = GetBinCutPatternRow(flowName, dataRow, row, sourceRow, pattern, intSkipTest);
                            binCutPatternRow.SheetName = dataRow.BinCutInstanceRow.SheetName;
                            if (!string.IsNullOrEmpty(pattern.Trim()))
                            {
                                binCutPatternRows.Add(binCutPatternRow);
                            }
                        }
                    }
                    else
                    {
                        foreach (string pattern in dataRow.PatternList)
                        {
                            BinCutPatternRow binCutPatternRow = GetBinCutPatternRow(flowName, dataRow, row, sourceRow, pattern, intSkipTest);
                            binCutPatternRow.SheetName = dataRow.BinCutInstanceRow.SheetName;
                            if (!string.IsNullOrEmpty(pattern.Trim()))
                            {
                                binCutPatternRows.Add(binCutPatternRow);
                            }
                        }
                    }

                    if (!binCutPatternRows.Any())
                    {
                        foreach (string pattern in dataRow.PatternList)
                        {
                            var binCutPatternRow = new BinCutPatternRow(flowName)
                            {
                                RowNum = dataRow.BinCutInstanceRow.RowNum,
                                BinCutInstanceRow = dataRow.BinCutInstanceRow,
                                Voltage = Type,
                                BinningDomain = sourceRow.BinningDomain,
                                PerformanceMode = sourceRow.PerformanceMode,
                                Type = sourceRow.ColumnName.ToString(),
                                Condition = sourceRow.ColumnContent,
                                DeviceCondition = GetDeviceCondition(dataRow),
                                InstanceName = row.TestName,
                                Pattern = pattern,
                                PattrenSetName = string.IsNullOrEmpty(dataRow.PatSetName) ? dataRow.PayloadList[0] : dataRow.PatSetName,
                                Jobs = row.FinalJobs,
                                Instance = dataRow.BinCutInstanceRow.Instance,
                                IsInterpolateSkip = intSkipTest
                            };
                            if (!string.IsNullOrEmpty(pattern.Trim()))
                            {
                                binCutPatternRow.SheetName = dataRow.BinCutInstanceRow.SheetName;
                                binCutPatternRows.Add(binCutPatternRow);
                            }
                        }
                    }
                }
            }
            return instanceRowList;
        }

        private List<FlowRow> GenFlowLists(List<BinCutRowForSort> binCutRows, bool isPost, bool isCsharp = false)
        {
            var flowRows = new List<FlowRow>();
            for (int index = 0; index < binCutRows.Count; index++)
            {
                BinCutRowForSort binCutRow = binCutRows[index];
                BinCutSourceItem sourceRow = binCutRow.BinCutSourceRow;
                BinCutFinalInstanceRow dataRow = binCutRow.BinCutFinalInstanceRow;
                bool isTmps = sourceRow.ColumnContent.ContainsIgnoreCase("Flow_TMPS");
                var binCutNonBinCutInstance = new BinCutNonBinCutInstance();

                #region body
                InstanceInterface = new BinCutInstanceInterfaceFactory().GetBinCutInstanceInterface(dataRow, sourceRow, BinCutInputManager, isPost);
                bool isHvcc = Type == "HBV";
                FlowRow testFlowRow = InstanceInterface.GenerateFlowRow(isHvcc || isPost, isTmps, isCsharp);
                #endregion

                dataRow.BinCutInstanceRow.SiteVar = sourceRow.ColumnName.Equals(EnumColumnName.E1Voltage) ? sourceRow.Enable : dataRow.BinCutInstanceRow.SiteVar;

                List<FlowRow> binTables = GenBinTable(isPost, testFlowRow, binCutRow, isCsharp);

                List<FlowRow> ifFlowRows = binCutNonBinCutInstance.GetIfFlowRows(dataRow, [testFlowRow], binTables);
                flowRows.AddRange(ifFlowRows);
            }
            return flowRows;
        }

        public SubFlowSheet GenerateFlowRows(string sheetName, List<BinCutRowForSort> binCutRows, bool isPost, bool isCsharp = false)
        {

            var flowSheet = new SubFlowSheet(sheetName);
            //flowSheet.FlowRows.Add(NwireSingleton.Instance().SettingInfo.GetNwireCall(sheetName));
            flowSheet.AddStartRows();
            flowSheet.Rows.AddRange(GenFlowLists(binCutRows, isPost, isCsharp));
            flowSheet.AddEndRows();
            return flowSheet;
        }
        private List<string> GetExistSheetNames()
        {
            var existSheetNames = new List<string>();
            existSheetNames.AddRange(BinCutInputManager.BinCutInstanceSheets.Where(x => x.PmOrderColNumber != -1 && x.PerfModeColNumber != -1).Select(x => x.SheetName));
            existSheetNames.AddRange(BinCutInputManager.BinCutInstanceSheets.Where(x => x.PmOrderColNumber == -1 || x.PerfModeColNumber == -1).Select(x => x.SheetName));
            return existSheetNames;
        }

        private bool GetIsOrderByPmOrder(List<BinCutRowForSort> binCutRows)
        {
            var inputSheetNames = binCutRows.Select(x => x.BinCutFinalInstanceRow.BinCutInstanceRow.SheetName).Distinct().ToList();
            var perModes = binCutRows.Select(x => x.BinCutSourceRow.PerformanceMode).Distinct().ToList();
            foreach (string inputSheetName in inputSheetNames)
            {
                var binCutInstanceSheets = BinCutInputManager.BinCutInstanceSheets.Where(x => x.SheetName.Equals(inputSheetName, StringComparison.CurrentCultureIgnoreCase)).ToList();
                if (binCutInstanceSheets.Count != 0)
                {
                    foreach (BinCutInstanceSheet binCutInstanceSheet in binCutInstanceSheets)
                    {
                        if (binCutInstanceSheet.PmOrderColNumber != -1 && binCutInstanceSheet.PerfModeColNumber != -1 &&
                            binCutInstanceSheet.Rows.Where(x => !string.IsNullOrEmpty(x.PerfMode) && perModes.Exists(y => y.Equals(x.PerfMode, StringComparison.OrdinalIgnoreCase)) && x.PmOrder != -1.0).ToList().Any())
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        protected virtual List<FlowRow> GenBinTable(bool isPost, FlowRow row, BinCutRowForSort binCutRow, bool isCsharp = false)
        {
            var binTables = new List<FlowRow>();
            if (isCsharp)
            {
                FlowRow flowRow = InstanceInterface.GetBinTableRowBv();
                if (binCutRow.BinCutSourceRow.ColumnName.Equals(EnumColumnName.FUNC))
                {
                    flowRow.Env = "UnknownTestType";
                }

                if (!binCutRow.BinCutSourceRow.InstOrCallFlowByBms)
                {
                    binTables.Add(string.IsNullOrEmpty(flowRow.Parameter) ? AddFailStopBinTable() : flowRow);
                }
            }
            else
            {
                binTables.Add(AddFailStopBinTable());
            }

            return binTables;
        }

        protected FlowRow AddFailStopBinTable()
        {
            var flowRow = new FlowRow { Opcode = OpCode.BinTable, Parameter = BinCutConstant.VddBinningFailStop };
            return flowRow;
        }

        private BinCutPatternRow GetBinCutPatternRow(string flowName, BinCutFinalInstanceRow dataRow, InstanceRow row, BinCutSourceItem sourceRow, string pattern, bool intSkipTest)
        {
            var binCutPatternRow = new BinCutPatternRow(flowName)
            {
                RowNum = dataRow.BinCutInstanceRow.RowNum,
                BinCutInstanceRow = dataRow.BinCutInstanceRow,
                Voltage = Type,
                BinningDomain = sourceRow.BinningDomain,
                PerformanceMode = sourceRow.PerformanceMode,
                Type = sourceRow.ColumnName.ToString(),
                Condition = sourceRow.ColumnContent,
                DeviceCondition = GetDeviceCondition(dataRow),
                InstanceName = row.TestName,
                PattrenSetName = dataRow.PayloadList.Count == 1 && sourceRow.ColumnName == EnumColumnName.Mbist
                    ? dataRow.PayloadList[0] : dataRow.PatSetName,
                Jobs = row.FinalJobs,
                Pattern = pattern,
                Instance = dataRow.BinCutInstanceRow.Instance,
                IsInterpolateSkip = intSkipTest
            };
            return binCutPatternRow;
        }

        private List<InstanceRow> GetInstanceRow(BinCutFinalInstanceRow dataRow)
        {
            var rows = new List<InstanceRow>();
            if (dataRow.BinCutInstanceRow.Type == BincutInstanceType.Hardip)
            {
                rows.AddRange(InstanceInterface.GenerateInstanceByTestName());
            }
            else
            {
                rows.Add(InstanceInterface.GenerateInstance());

            }
            return rows;
        }

        private void GenBlankRow(List<BinCutRow> instanceNameWithJobs)
        {
            foreach (BinCutRow instanceNameWithJob in instanceNameWithJobs)
            {
                if (instanceNameWithJob.BinCutFinalInstanceRows.Count == 0)
                {
                    instanceNameWithJob.BinCutFinalInstanceRows.Add(instanceNameWithJob.BinCutSourceItem.FillBlankRow());
                }
            }
        }

        private string GetDeviceCondition(BinCutFinalInstanceRow dataRow)
        {
            if (dataRow.BinCutInstanceRow != null)
            {
                if (!string.IsNullOrEmpty(dataRow.BinCutInstanceRow.SiteVar))
                {
                    return dataRow.BinCutInstanceRow.SiteVar;
                }

                if (!string.IsNullOrEmpty(dataRow.BinCutInstanceRow.EnableAndDevice))
                {
                    return dataRow.BinCutInstanceRow.EnableAndDevice;
                }
            }
            return "";
        }
    }
}
