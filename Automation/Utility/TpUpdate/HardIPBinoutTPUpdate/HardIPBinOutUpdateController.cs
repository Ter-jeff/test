using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using FlowSheet = IgxlLib.IgxlSheets.SubFlowSheet;

namespace Automation.Utility.TpUpdate.HardIPBinoutTPUpdate
{
    public class HardIpBinOutUpdateController
    {
        private readonly Dictionary<string, string> _typeDic = new Dictionary<string, string> {
            { "test", "Function" },
            { "use-limit", "Parametric" }
        };
        private readonly bool _isCommandLine;
        private Dictionary<string, HashSet<string>> _failActionDic = new Dictionary<string, HashSet<string>>();
        private HashSet<string> _binTableUpdateCached = new HashSet<string>();
        private const string ConIvdm = "IVDM";
        private const string ConIvdm3X = "3X";

        #region Regular Expression
        private readonly Regex _regFlowPrefix = new Regex("Flow_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regBinOut = new Regex("^binout", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regNoBinOut = new Regex(@"non*\s*bin\s*out", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regSkipBinOut = new Regex(@"^skipbintable", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regRenameFlag = new Regex(@"(?<Block>\w+)_(?<SubBlock>\w+)_(?<pattern>[(PP)(MN)(CZ)(DD)(DP)]\w+)_(?<voltage>[NHL]V)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regHacTestName = new Regex("^HAC_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regCPJobs = new Regex(@"^CP", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regFTJobs = new Regex(@"^FT|^WLFT", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Regex _regDeferFlowHip = new Regex(@"(?=.*Defer)(?=.*(HardIp|Hip))", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private const string SkipBinTableEnable = "SkipBinTable";
        #endregion

        public HardIpBinOutUpdateController(bool isCommandLine)
        {
            _isCommandLine = isCommandLine;
        }

        /// <summary>
        /// Reference the original status and high/low limits if update column is empty
        /// </summary>
        /// <param name="binoutSheet">Binout sheet from binout report</param>
        /// <returns></returns>
        public void ReferenceStatusAndLimits(BinOutStatusHipListSheet binoutSheet, bool isOverWriteLimits = true)
        {
            foreach (BinOutStatusHipListRow row in binoutSheet.Rows)
            {
                // Filling original binout status to blank updated column
                foreach (string jobPart in row.StatusDictionary.Keys)
                {
                    string job = jobPart.Split('_')[0] + "_";
                    string originalStatus = row.StatusDictionary[jobPart];

                    if (row.BinOutEnableDictionary.ContainsKey(job) && string.IsNullOrEmpty(row.BinOutEnableDictionary[job]))
                    {
                        row.BinOutEnableDictionary[job] = originalStatus;
                    }

                    if (isOverWriteLimits)
                    {
                        // Filling original limits to blank updated limit column
                        RefernceOrgLimits(row.HiLimitDic, row.UpdatedHiLimitDic, job);
                        RefernceOrgLimits(row.LoLimitDic, row.UpdatedLoLimitDic, job);
                    }
                }
            }
        }

        /// <summary>
        /// Access the sheets for hardip and defer bin flow
        /// </summary>
        /// <param name="blocks"></param>
        /// <param name="subFlowSheets"></param>
        /// <returns></returns>
        public List<string> AccessFileList(List<string> blocks, Dictionary<string, FlowSheet> subFlowSheets)
        {
            var files = subFlowSheets.Keys.ToList();

            //Only Access Flow Sheet By Header specified at top parameter
            var fileNames = files.Where(
                p => _regFlowPrefix.IsMatch(p) && (
                        blocks.Exists(q => Regex.IsMatch(Path.GetFileNameWithoutExtension(p), q, RegexOptions.IgnoreCase)) ||
                        blocks.Exists(q => Regex.IsMatch(GetBlockFromSheetName(p), q, RegexOptions.IgnoreCase))) ||
                        _regDeferFlowHip.IsMatch(Path.GetFileNameWithoutExtension(p))
                    ).ToList();

            return fileNames;
        }

        /// <summary>
        /// Get hardip block name from sheet name
        /// </summary>
        /// <param name="sheetpath"></param>
        /// <returns></returns>
        public string GetBlockFromSheetName(string sheetpath)
        {
            string sheetname = Path.GetFileNameWithoutExtension(sheetpath) ?? sheetpath;

            if (sheetname.StartsWith("Flow_Hardip_", StringComparison.CurrentCultureIgnoreCase))
            {
                string subsheetname = sheetname.Substring(12);
                return subsheetname.Replace("_", "");
            }
            return sheetpath;
        }

        public Dictionary<string, List<string>> GetHardipBintableRows(List<BinTableSheet> binTableSheets)
        {
            var hardipBinItems = new Dictionary<string, List<string>>();
            BinTableSheet bintable = GetHardipBintableFromProgram(binTableSheets);
            if (bintable != null)
            {
                var binItems = bintable.Rows.GroupBy(x => x.Name)
                    .ToDictionary(x => x.Key, x => x.Where(y => !string.IsNullOrEmpty(y.ItemList)).Select(z => z.ItemList).Distinct().ToList());
                foreach (KeyValuePair<string, List<string>> binItem in binItems)
                {
                    var flags = new List<string>();
                    foreach (string flag in binItem.Value)
                    {
                        string[] arr = flag.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        flags.AddRange(arr);
                    }
                    if (flags.Any())
                    {
                        flags = flags.Distinct().ToList();
                        if (hardipBinItems.ContainsKey(binItem.Key))
                        {
                            hardipBinItems[binItem.Key].AddRange(flags);
                        }
                        else
                        {
                            hardipBinItems.Add(binItem.Key, flags);
                        }
                    }
                }
            }
            else if (_isCommandLine)
            {
                Response.Report("[Missing] Bin_Table_HardIP sheet can't be found !!!", EnumMessageLevel.Warning);
            }

            return hardipBinItems;
        }

        internal BinTableSheet GetHardipBintableFromProgram(List<BinTableSheet> binTableSheets)
        {
            if (binTableSheets.Any())
            {
                return binTableSheets.Find(
                        x => x.Name.Equals("Bin_Table_HardIP", StringComparison.CurrentCultureIgnoreCase) &&
                             x.Rows.Count > 0
                    );
            }
            return null;
        }

        public List<FlowSheet> UpdateFlowRowsMainByMultiJobs(List<FlowSheet> sheets, Dictionary<string, List<KeyValuePair<string, List<BinOutStatusHipListRow>>>> binoutByBlocks, Dictionary<string, List<string>> bintable, HashSet<string> currentJobs)
        {
            var resultFlowSheet = new Dictionary<string, FlowSheet>();

            foreach (KeyValuePair<string, List<KeyValuePair<string, List<BinOutStatusHipListRow>>>> binoutByBlock in binoutByBlocks)
            {
                string blockName = binoutByBlock.Key;

                if (_isCommandLine)
                {
                    Response.Report($"Process block : {blockName}", EnumMessageLevel.Info);
                }

                var allFlowItems = new List<string>();
                var funcItemJobPart = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (FlowSheet sheet in sheets)
                {
                    var allbinoutItems = new Dictionary<string, List<BinOutStatusHipListRow>>();
                    string temp = "";

                    #region update test and use-limit items
                    _failActionDic = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                    for (int testIdx = 0; testIdx < sheet.Rows.Count; testIdx++)
                    {
                        FlowRow flowRow = sheet.Rows[testIdx];
                        if (!_typeDic.ContainsKey(flowRow.Opcode.ToLower()))
                        {
                            continue;
                        }

                        string testType = _typeDic[flowRow.Opcode.ToLower()];
                        string testItem = flowRow.Parameter + "#" + testType;
                        allFlowItems.Add(testItem);
                        if (!string.IsNullOrEmpty(temp) && temp.Equals(testItem))
                        {
                            continue;
                        }

                        temp = testItem;
                        var binoutItems = binoutByBlock.Value.Where(x => x.Key.Equals(testItem, StringComparison.OrdinalIgnoreCase)).ToList();
                        if (binoutItems.Any())
                        {
                            var flowAllTestItems = sheet.Rows.Where(x => x.Parameter.Equals(flowRow.Parameter, StringComparison.CurrentCultureIgnoreCase)).ToList();
                            bool isFuncItem = testType == "Function";
                            if (isFuncItem)
                            {
                                funcItemJobPart.Clear();
                            }
                            else
                            {
                                foreach (BinOutStatusHipListRow binoutItem in binoutItems.First().Value)
                                {
                                    binoutItem.IsUpdated.Clear();
                                }
                            }

                            HashSet<string> jobparts = GetTestJobs(binoutItems.First());
                            PrintNoTestItemToErrorReport(binoutItems, jobparts);
                            bool canUpdateBinoutStatus = true;
                            var paramIsUpdateRec = new Dictionary<string, HashSet<string>>();

                            foreach (string jobpart in jobparts)
                            {
                                ProcessJobPartForFlowRow(jobpart, flowRow, testIdx, sheet, binoutItems, currentJobs, jobparts, isFuncItem, funcItemJobPart, flowAllTestItems, paramIsUpdateRec, ref canUpdateBinoutStatus);
                            }

                            if (canUpdateBinoutStatus)
                            {
                                SaveBinoutItemsForUpdated(ref allbinoutItems, binoutItems, flowRow.Parameter);
                            }
                        }
                    }
                    #endregion

                    #region update bintable
                    List<FlowSheet> binTableUpdateSheets = UpdateBinTableForSheet(sheet, sheets, allbinoutItems, bintable);
                    #endregion

                    if (allbinoutItems.Count > 0)
                    {
                        // Add current updated sheet
                        resultFlowSheet.TryAdd(sheet.Name, sheet);

                        // Add flow sheet for updating the bintable rows
                        // ex. Flow_Defer, or other sheets with _1, _2 ...
                        foreach (FlowSheet binTableRowUpdateSheet in binTableUpdateSheets)
                        {
                            resultFlowSheet.TryAdd(binTableRowUpdateSheet.Name, binTableRowUpdateSheet);
                        }
                    }
                }

                if (allFlowItems.Any())
                {
                    ReportMissingInstances(allFlowItems, binoutByBlock);
                }
            }

            return resultFlowSheet.Values.ToList();
        }

        private void ReportMissingInstances(List<string> allFlowItems, KeyValuePair<string, List<KeyValuePair<string, List<BinOutStatusHipListRow>>>> binoutByBlock)
        {
            foreach (KeyValuePair<string, List<BinOutStatusHipListRow>> binoutItem in binoutByBlock.Value)
            {
                if (binoutItem.Value.Any() &&
                    !allFlowItems.Any(x => x.Equals(binoutItem.Key, StringComparison.CurrentCultureIgnoreCase)))
                {
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MissingInstance_02, binoutItem.Value.First().SourceSheetName, binoutItem.Value.First().RowNum, 0, [binoutItem.Key]);
                }
            }
        }

        /// <summary>
        /// Currently, there is no job insertion requirement for the HardIP flow bin table.
        /// The flow bin table is only updated when the status is set to `SkipBinTable` on the BOR.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sheets"></param>
        /// <param name="allbinoutItems"></param>
        /// <param name="bintable"></param>
        /// <returns></returns>
        private List<FlowSheet> UpdateBinTableForSheet(FlowSheet sheet, List<FlowSheet> sheets, Dictionary<string, List<BinOutStatusHipListRow>> allbinoutItems, Dictionary<string, List<string>> bintable)
        {
            var binTableUpdateSheets = new List<FlowSheet>();

            foreach (KeyValuePair<string, List<BinOutStatusHipListRow>> item in allbinoutItems)
            {
                HashSet<string> jobs = GetTestJobs(item);
                BinOutStatusHipListRow binoutreport = item.Value.First();
                if (!_failActionDic.ContainsKey(item.Key))
                {
                    continue;
                }

                bool isSkipBinTable = item.Value.Any(x => x.BinOutEnableDictionary.Values.Any(_regSkipBinOut.IsMatch));
                if (!isSkipBinTable)
                {
                    continue;
                }

                HashSet<string> failflags = _failActionDic[item.Key];
                if (!failflags.Any())
                {
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MissingFlag_01, binoutreport.SourceSheetName, binoutreport.RowNum, 0, [item.Key]);
                    continue;
                }
                string currBinName = GetBinNameByFailFalg(bintable, failflags);
                if (string.IsNullOrEmpty(currBinName))
                {
                    string flagName = string.Join(",", failflags);
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MissingBinTable_01, binoutreport.SourceSheetName, binoutreport.RowNum, 0, [flagName]);
                    continue;
                }

                // skip if the corresponding flow bin table has been updated to `SkipBinTable`
                if (_binTableUpdateCached.Contains(currBinName))
                {
                    continue;
                }

                var flowBintables = sheets.SelectMany(consheet =>
                            consheet.Rows.Where(x =>
                                x.Parameter.Equals(currBinName, StringComparison.CurrentCultureIgnoreCase) &&
                                x.Opcode.Equals(OpCode.BinTable, StringComparison.CurrentCultureIgnoreCase))
                            .Select(row => (UpdateFlowSheet: consheet, RowData: row))
                        ).ToList();

                if (!flowBintables.Any())
                {
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MissingBinTable_02, binoutreport.SourceSheetName, 0, 0, [currBinName, sheet.Name]);
                    continue;
                }

                // Handle for skip bintable case that gating with an addtional enable word `SkipBinTable`
                _binTableUpdateCached.Add(currBinName);
                foreach ((FlowSheet flowSheet, FlowRow flowBintable) in flowBintables)
                {
                    binTableUpdateSheets.Add(flowSheet);
                    flowBintable.Enable = string.IsNullOrEmpty(flowBintable.Enable) ? SkipBinTableEnable :
                        $"({flowBintable.Enable})&&({SkipBinTableEnable})";
                }
            }

            return binTableUpdateSheets;
        }

        private void ProcessJobPartForFlowRow(string jobpart, FlowRow flowRow, int testIdx, FlowSheet sheet, List<KeyValuePair<string, List<BinOutStatusHipListRow>>> binoutItems, HashSet<string> currentJobs, HashSet<string> jobparts, bool isFuncItem, HashSet<string> funcItemJobPart, List<FlowRow> flowAllTestItems, Dictionary<string, HashSet<string>> paramIsUpdateRec, ref bool canUpdateBinoutStatus)
        {
            string[] jobpartarr = jobpart.Split('_');
            string job = jobpartarr[0];
            string part = jobpartarr.Length > 1 ? jobpartarr[1] : "";

            FlowRow row = flowRow;
            List<FlowRow> flowItems = GetItems(sheet.Rows, row, testIdx, job, part);
            var newbinoutItems = binoutItems.First().Value.Where(item => item.EnableDictionary.ContainsKey(jobpart)).ToList();
            if (!flowItems.Any())
            {
                return;
            }

            List<FlowRow> measureItems = isFuncItem ? GetUseLimitByFunctionTest(job, part, row, sheet.Rows, flowItems.First()) : new List<FlowRow>();
            bool allRowsUpdateBinoutStatusIsSame = false;
            bool allRowsUpdateLimitValuesIsSame = false;
            bool countIsSame = flowItems.Count == newbinoutItems.Count;

            if (!countIsSame)
            {
                if (newbinoutItems.Select(x => x.BinOutEnableDictionary[jobpart]).Distinct().Count() == 1)
                {
                    allRowsUpdateBinoutStatusIsSame = true;
                }

                if (newbinoutItems.Select(x => x.UpdatedLoLimitDic[jobpart]).Distinct().Count() == 1 &&
                    newbinoutItems.Select(x => x.UpdatedHiLimitDic[jobpart]).Distinct().Count() == 1)
                {
                    allRowsUpdateLimitValuesIsSame = true;
                }

                if (!allRowsUpdateBinoutStatusIsSame && !allRowsUpdateLimitValuesIsSame)
                {
                    canUpdateBinoutStatus = false;
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MismatchCount_01, binoutItems.First().Value.First().SourceSheetName, binoutItems.First().Value.First().RowNum, 0,
                        [jobpart, binoutItems.First().Value.Count.ToString(), sheet.Name, flowItems.Count.ToString()]);
                }
                if (!allRowsUpdateBinoutStatusIsSame && allRowsUpdateLimitValuesIsSame)
                {
                    canUpdateBinoutStatus = false;
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MismatchCount_02, binoutItems.First().Value.First().SourceSheetName, binoutItems.First().Value.First().RowNum, 0,
                        [jobpart, binoutItems.First().Value.Count.ToString(), sheet.Name, flowItems.Count.ToString()]);
                }
                else if (allRowsUpdateBinoutStatusIsSame && !allRowsUpdateLimitValuesIsSame)
                {
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MismatchCount_03, binoutItems.First().Value.First().SourceSheetName, binoutItems.First().Value.First().RowNum, 0,
                        [jobpart, binoutItems.First().Value.Count.ToString(), sheet.Name, flowItems.Count.ToString()]);
                }
            }

            if (countIsSame)
            {
                List<BinOutStatusHipListRow> finalBinoutItems = newbinoutItems;
                for (int i = 0; i < flowItems.Count; i++)
                {
                    FlowRow flowItem = flowItems[i];
                    BinOutStatusHipListRow binoutItem = finalBinoutItems[i];
                    ProcessUpdateTp(binoutItem, currentJobs, jobparts, jobpart, isFuncItem, funcItemJobPart, sheet, flowItem, flowAllTestItems, measureItems);
                }
            }
            else if (isFuncItem)
            {
                ErrorReportManager.AddError(BinOutReportErrorType.E_MismatchCount_04, binoutItems.First().Value.First().SourceSheetName, binoutItems.First().Value.First().RowNum, 0,
                    [flowItems.First().Parameter, jobpart]);
            }
            // update the flag and limits by test name
            else
            {
                UpdateFlowByTestName(flowItems, newbinoutItems, binoutItems, currentJobs, jobparts, jobpart, isFuncItem, funcItemJobPart, sheet, flowAllTestItems, measureItems, paramIsUpdateRec);
            }
        }

        private void UpdateFlowByTestName(List<FlowRow> flowItems, List<BinOutStatusHipListRow> newbinoutItems, List<KeyValuePair<string, List<BinOutStatusHipListRow>>> binoutItems, HashSet<string> currentJobs, HashSet<string> jobparts, string jobpart, bool isFuncItem, HashSet<string> funcItemJobPart, FlowSheet sheet, List<FlowRow> flowAllTestItems, List<FlowRow> measureItems, Dictionary<string, HashSet<string>> paramIsUpdateRec)
        {
            Dictionary<string, List<FlowRow>> finalFlowItems = GetFlowItemByTestName(flowItems);
            Dictionary<string, List<BinOutStatusHipListRow>> finalBinoutItems = GetBinoutItemsByTestName(newbinoutItems, jobpart);

            foreach (string testName in finalFlowItems.Keys)
            {
                if (!finalBinoutItems.TryGetValue(testName, out List<BinOutStatusHipListRow> item))
                {
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MismatchCount_05, binoutItems.First().Value.First().SourceSheetName, binoutItems.First().Value.First().RowNum, 0,
                        [flowItems.First().Parameter, finalFlowItems[testName].First().TName, jobpart]);
                    continue;
                }

                bool isCountSame = finalFlowItems[testName].Count == item.Count;
                bool isStatusSame = finalBinoutItems[testName].Select(x => x.BinOutEnableDictionary[jobpart]).Distinct().Count() == 1;
                bool isUpdateLimitSame = finalBinoutItems[testName].Select(x => x.UpdatedLoLimitDic[jobpart]).Distinct().Count() == 1 && finalBinoutItems[testName].Select(x => x.UpdatedHiLimitDic[jobpart]).Distinct().Count() == 1;

                if (!isCountSame && (!isStatusSame || !isUpdateLimitSame))
                {
                    ErrorReportManager.AddError(BinOutReportErrorType.E_MismatchCount_06, binoutItems.First().Value.First().SourceSheetName, binoutItems.First().Value.First().RowNum, 0,
                        [flowItems.First().Parameter, finalFlowItems[testName].First().TName, jobpart]);
                    continue;
                }

                for (int i = 0; i < finalFlowItems[testName].Count; i++)
                {
                    FlowRow flowItem = finalFlowItems[testName][i];
                    BinOutStatusHipListRow binoutItem = !isCountSame ? finalBinoutItems[testName].First() : finalBinoutItems[testName][i];

                    ProcessUpdateTp(binoutItem, currentJobs, jobparts, jobpart, isFuncItem, funcItemJobPart, sheet, flowItem, flowAllTestItems, measureItems
                    );

                    if (!isCountSame)
                    {
                        if (!paramIsUpdateRec.ContainsKey(testName))
                        {
                            paramIsUpdateRec[testName] = binoutItem.IsUpdated.ToHashSet();
                        }
                        else
                        {
                            paramIsUpdateRec[testName].AddRange(binoutItem.IsUpdated);
                        }

                        binoutItem.IsUpdated.Clear();
                    }
                }

                if (!isCountSame)
                {
                    finalBinoutItems[testName].First().IsUpdated = paramIsUpdateRec[testName].ToHashSet();
                }
            }
        }

        internal HashSet<string> GetTestJobs(KeyValuePair<string, List<BinOutStatusHipListRow>> binoutItems)
        {
            var jobs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (binoutItems.Value.First().EnableDictionary.Count > 1)
            {
                foreach (BinOutStatusHipListRow item in binoutItems.Value)
                {
                    jobs.AddRange(item.EnableDictionary.Select(job =>
                    {
                        string[] jobArr = job.Key.Split('_');
                        string jobpart = string.Join("_", jobArr[0], jobArr[1]);
                        return item.BinOutEnableDictionary.ContainsKey(jobpart) ? jobpart : "";
                    }));
                }
                jobs = jobs.Where(x => !string.IsNullOrEmpty(x)).ToHashSet();
            }
            else
            {
                jobs.Add(binoutItems.Value.First().BinOutEnableDictionary.First().Key);
            }

            return jobs;
        }

        private void PrintNoTestItemToErrorReport(IEnumerable<KeyValuePair<string, List<BinOutStatusHipListRow>>> binoutItems, HashSet<string> testJobParts)
        {
            foreach (KeyValuePair<string, List<BinOutStatusHipListRow>> items in binoutItems)
            {
                foreach (BinOutStatusHipListRow item in items.Value)
                {
                    string noTestJob = "";
                    if (HasNoTestItemNeedToUpdate(item, testJobParts, ref noTestJob))
                    {
                        ErrorReportManager.AddError(BinOutReportErrorType.W_MismatchJob_01, item.SourceSheetName, item.RowNum, 0, [item.Testinstance, noTestJob]);
                    }
                }
            }
        }

        internal bool HasNoTestItemNeedToUpdate(BinOutStatusHipListRow item, HashSet<string> testJobParts, ref string noTestJob)
        {
            foreach (KeyValuePair<string, string> state in item.BinOutEnableDictionary)
            {
                if (!testJobParts.Contains(state.Key) && !string.IsNullOrEmpty(state.Value))
                {
                    noTestJob = state.Key;
                    return true;
                }
            }

            foreach (KeyValuePair<string, string> loLimit in item.UpdatedLoLimitDic)
            {
                if (!testJobParts.Contains(loLimit.Key) && !string.IsNullOrEmpty(loLimit.Value))
                {
                    noTestJob = loLimit.Key;
                    return true;
                }
            }

            foreach (KeyValuePair<string, string> hiLimit in item.UpdatedHiLimitDic)
            {
                if (!testJobParts.Contains(hiLimit.Key) && !string.IsNullOrEmpty(hiLimit.Value))
                {
                    noTestJob = hiLimit.Key;
                    return true;
                }
            }

            return false;
        }

        private List<FlowRow> GetItems(List<FlowRow> sourceFlowRows, FlowRow currRow, int currIndex, string currectJob, string currectPart)
        {
            var items = new List<FlowRow>();
            List<FlowRow> newFlowRows = sourceFlowRows.GetRange(currIndex, sourceFlowRows.Count - currIndex);
            foreach (FlowRow row in newFlowRows)
            {
                if (!row.Opcode.Equals(currRow.Opcode, StringComparison.CurrentCultureIgnoreCase))
                {
                    break;
                }

                if (row.Parameter.Equals(currRow.Parameter, StringComparison.CurrentCultureIgnoreCase) &&
                    row.Opcode.Equals(currRow.Opcode, StringComparison.CurrentCultureIgnoreCase) &&
                    CheckJobPartIsMatch(row.Job, currectJob) &&
                    CheckJobPartIsMatch(row.Part, currectPart))
                {
                    items.Add(row);
                }
            }
            return items;
        }

        private List<FlowRow> GetUseLimitByFunctionTest(string currJob, string currPart, FlowRow currRow, List<FlowRow> sourceFlowRows, FlowRow firstFlowItem)
        {
            var measureItems = new List<FlowRow>();

            int startIndex = sourceFlowRows.FindIndex(x => x.Equals(firstFlowItem));
            List<FlowRow> newFlowRows = sourceFlowRows.GetRange(startIndex + 1, sourceFlowRows.Count - (startIndex + 1));
            foreach (FlowRow row in newFlowRows)
            {
                if (row.Opcode.Equals(OpCode.Test, StringComparison.CurrentCultureIgnoreCase))
                {
                    break;
                }

                if (row.Parameter.Equals(currRow.Parameter, StringComparison.CurrentCultureIgnoreCase) &&
                    row.Opcode.Equals(OpCode.UseLimit, StringComparison.CurrentCultureIgnoreCase) &&
                    (row.Job.ContainsIgnoreCase(currJob.ToUpper()) || string.IsNullOrEmpty(row.Job)) &&
                    (row.Part.ContainsIgnoreCase(currPart.ToUpper()) || string.IsNullOrEmpty(row.Part))
                   )
                {
                    measureItems.Add(row);
                }

            }

            return measureItems;
        }

        private HashSet<string> GetNeedUpdateTpJob(BinOutStatusHipListRow binoutItem, HashSet<string> currentJobs)
        {
            var jobs = currentJobs.Where(job => !binoutItem.IsUpdated.Contains(job) && binoutItem.BinOutEnableDictionary.ContainsKey(job) &&
                                                 (_regNoBinOut.IsMatch(binoutItem.BinOutEnableDictionary[job]) ||
                                                 _regBinOut.IsMatch(binoutItem.BinOutEnableDictionary[job]) ||
                                                 _regSkipBinOut.IsMatch(binoutItem.BinOutEnableDictionary[job]) ||
                                                 !string.IsNullOrEmpty(binoutItem.UpdatedHiLimitDic[job]))).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return jobs;
        }

        private HashSet<string> UpdateValueIsSame(BinOutStatusHipListRow binoutItem, string currJobPart, HashSet<string> needUpdateJobs)
        {
            var updateSameValueJobPart = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { currJobPart };
            string currBinoutState = binoutItem.BinOutEnableDictionary[currJobPart];
            bool isCurrStateUpdateFlag = _regBinOut.IsMatch(currBinoutState) || _regSkipBinOut.IsMatch(currBinoutState);

            string currLolimit = binoutItem.UpdatedLoLimitDic[currJobPart];
            string currHilimit = binoutItem.UpdatedHiLimitDic[currJobPart];

            foreach (string jobPart in binoutItem.BinOutEnableDictionary.Keys)
            {
                if (jobPart == currJobPart)
                {
                    continue;
                }

                if (!needUpdateJobs.Contains(jobPart))
                {
                    continue;
                }

                string otherBinoutState = binoutItem.BinOutEnableDictionary[jobPart];
                bool isOtherStateUpdateFlag = _regBinOut.IsMatch(otherBinoutState) || _regSkipBinOut.IsMatch(otherBinoutState);

                string otherLolimit = binoutItem.UpdatedLoLimitDic[jobPart];
                string otherHilimit = binoutItem.UpdatedHiLimitDic[jobPart];

                // Status `BinOut` and `SkipBinTable` will update the fail flag back => treat as the same
                // Categorized the status to determine if update fail flag or not
                if (isCurrStateUpdateFlag != isOtherStateUpdateFlag ||
                    !currLolimit.Equals(otherLolimit, StringComparison.CurrentCultureIgnoreCase) ||
                    !currHilimit.Equals(otherHilimit, StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                // collect the job part with current job part if the updated values are the same
                updateSameValueJobPart.Add(jobPart);
            }

            return updateSameValueJobPart;
        }

        private bool InsertItemToProgram(bool needUpdateLimitToFunc, FlowSheet sheet, int currRowIdx, FlowRow currRow, HashSet<string> updateJobParts, BinOutStatusHipListRow binoutItem, List<FlowRow> allSameTestItems, List<FlowRow> measureItems)
        {
            string refJob = updateJobParts.FirstOrDefault();
            string updateJobs = "";
            string updateParts = "";
            foreach (string jobPart in updateJobParts)
            {
                string[] jobpartarr = jobPart.Split('_');
                updateJobs = string.IsNullOrEmpty(updateJobs) ? jobpartarr[0] : string.Join(",", updateJobs, jobpartarr[0]);

                if (jobpartarr.Length > 1)
                {
                    updateParts = string.IsNullOrEmpty(updateParts) ? jobpartarr[1] : string.Join(",", updateParts, jobpartarr[1]);
                }
            }

            FlowRow copy = currRow.Copy();
            bool needUpdate = UpdateItemToProgram(ref copy, binoutItem, refJob, allSameTestItems, sheet.Name);

            // Check jobs, if original jobs are not same as updated jobs, it should be updated too
            if (!needUpdate)
            {
                var orgJobs = currRow.GetJobs().ToHashSet();
                needUpdate = !orgJobs.SetEquals(updateJobs.Split(',').ToHashSet());
            }

            if (needUpdate)
            {
                copy.Job = updateJobs; // currJob;
                copy.Part = updateParts; // currPart;
                sheet.InsertRow(currRowIdx, copy);
                if (needUpdateLimitToFunc)
                {
                    InsertMeasureItemToProgram(sheet, ++currRowIdx, measureItems);
                }
                return true;
            }
            return false;
        }

        internal bool UpdateItemToProgram(ref FlowRow flowItem, BinOutStatusHipListRow binoutItem, string jobpart, List<FlowRow> flowAllTestItems, string sheetName)
        {
            bool needUpdate = false;
            if (!string.IsNullOrEmpty(binoutItem.UpdatedHiLimitDic[jobpart]))
            {
                string loHimit = binoutItem.UpdatedHiLimitDic[jobpart].Equals("N/A", StringComparison.OrdinalIgnoreCase) ?
                    "" : binoutItem.UpdatedHiLimitDic[jobpart];
                if (!flowItem.HiLim.Equals(loHimit, StringComparison.CurrentCultureIgnoreCase))
                {
                    if (!string.IsNullOrEmpty(flowItem.HiLim) && !double.TryParse(flowItem.HiLim.TrimStart('='), out double _))
                    {
                        ErrorReportManager.AddError(BinOutReportErrorType.W_InvalidLimit_01, binoutItem.SourceSheetName, binoutItem.RowNum, 0, ["high", flowItem.HiLim, flowItem.SheetName]);
                    }
                    flowItem.HiLim = loHimit;
                    needUpdate = true;
                }
            }

            if (!string.IsNullOrEmpty(binoutItem.UpdatedLoLimitDic[jobpart]))
            {
                string loLimit = binoutItem.UpdatedLoLimitDic[jobpart].Equals("N/A", StringComparison.OrdinalIgnoreCase) ?
                    "" : binoutItem.UpdatedLoLimitDic[jobpart];
                if (!flowItem.LoLim.Equals(loLimit, StringComparison.CurrentCultureIgnoreCase))
                {
                    if (!string.IsNullOrEmpty(flowItem.LoLim) && !double.TryParse(flowItem.LoLim.TrimStart('='), out double _))
                    {
                        ErrorReportManager.AddError(BinOutReportErrorType.W_InvalidLimit_01, binoutItem.SourceSheetName, binoutItem.RowNum, 0, ["low", flowItem.LoLim, flowItem.SheetName]);
                    }
                    flowItem.LoLim = loLimit;
                    needUpdate = true;
                }
            }

            string binOutStatus = binoutItem.BinOutEnableDictionary[jobpart];
            if (!string.IsNullOrEmpty(binOutStatus))
            {
                if (_regNoBinOut.IsMatch(binOutStatus))
                {
                    if (UpdateNoBinOutInfo(flowItem, sheetName))
                    {
                        needUpdate = true;
                    }
                }
                else if (_regBinOut.IsMatch(binOutStatus) || _regSkipBinOut.IsMatch(binOutStatus))
                {
                    try
                    {
                        string flagName = "";
                        foreach (FlowRow item in flowAllTestItems)
                        {

                            if (!string.IsNullOrEmpty(item.FailAction.Trim('\t')))
                            {
                                flagName = item.FailAction.Trim('\t');
                                break;
                            }

                            string[] comments = item.Comment1.Split(new[] { '\t', ' ' });
                            if (comments.Length > 1 && !string.IsNullOrEmpty(comments[1]))
                            {
                                flagName = comments[1];
                                break;
                            }
                        }

                        if (string.IsNullOrEmpty(flagName))
                        {
                            return needUpdate;
                        }

                        UpdateFailActionDic(flowItem.Parameter, flagName);
                        if (!flowItem.FailAction.Equals(flagName, StringComparison.CurrentCultureIgnoreCase))
                        {
                            flowItem.FailAction = flagName;
                            needUpdate = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Response.Report(string.Format(ex.ToString()), EnumMessageLevel.Error);
                    }
                }
            }

            return needUpdate;
        }

        private Dictionary<string, List<FlowRow>> GetFlowItemByTestName(List<FlowRow> flowItems)
        {
            var flowItemsDict = new Dictionary<string, List<FlowRow>>(StringComparer.OrdinalIgnoreCase);

            foreach (FlowRow flowItem in flowItems)
            {
                string newTestName = ProcessTpTestName(flowItem.TName);

                if (!flowItemsDict.ContainsKey(newTestName))
                {
                    flowItemsDict.Add(newTestName, new List<FlowRow> { flowItem });
                }
                else
                {
                    flowItemsDict[newTestName].Add(flowItem);
                }
            }

            return flowItemsDict;
        }

        private Dictionary<string, List<BinOutStatusHipListRow>> GetBinoutItemsByTestName(List<BinOutStatusHipListRow> binoutItems, string jobpart)
        {
            var binoutItemsDict = new Dictionary<string, List<BinOutStatusHipListRow>>(StringComparer.OrdinalIgnoreCase);

            foreach (BinOutStatusHipListRow binoutItem in binoutItems)
            {
                if (!string.IsNullOrEmpty(binoutItem.PatternPin))
                {
                    bool isCpMergePin = binoutItem.PatternPin.Contains("CP=");
                    bool isFtMergePin = binoutItem.PatternPin.Contains("FT=");

                    // the parametric row is only for CP
                    if (isCpMergePin && !isFtMergePin && !_regCPJobs.IsMatch(jobpart))
                    {
                        continue;
                    }

                    // the parametric row is only for FT
                    if (!isCpMergePin && isFtMergePin && !_regFTJobs.IsMatch(jobpart))
                    {
                        continue;
                    }
                }

                string newTestName;
                if (_regHacTestName.IsMatch(binoutItem.Testname) && binoutItem.Testname.Split('_').Length > 6)
                {
                    string sixthSegment = binoutItem.Testname.Split('_')[6];
                    string seventhSegment = binoutItem.Testname.Split('_')[7];
                    newTestName = sixthSegment + "_" + seventhSegment;
                }
                else
                {
                    newTestName = ProcessTpTestName(binoutItem.Testname);
                }

                if (!binoutItemsDict.ContainsKey(newTestName))
                {
                    binoutItemsDict.Add(newTestName, new List<BinOutStatusHipListRow> { binoutItem });
                }
                else
                {
                    binoutItemsDict[newTestName].Add(binoutItem);
                }
            }

            return binoutItemsDict;
        }

        internal string ProcessTpTestName(string tpTestName)
        {
            string newTestName;
            if (tpTestName.Contains("__"))
            {
                string[] testNameArr = tpTestName.Split(new[] { "__" }, StringSplitOptions.None);
                newTestName = testNameArr[1].Replace("_", "-") + "_" + testNameArr[0].Replace("_", "-");
            }
            else
            {
                newTestName = tpTestName.Replace("_", "-") + "_x";
            }

            return newTestName;
        }

        internal void SaveBinoutItemsForUpdated(ref Dictionary<string, List<BinOutStatusHipListRow>> allBinoutItems, IEnumerable<KeyValuePair<string, List<BinOutStatusHipListRow>>> binoutItems, string currFlowInstName)
        {
            foreach (KeyValuePair<string, List<BinOutStatusHipListRow>> binoutValue in binoutItems)
            {
                if (allBinoutItems.ContainsKey(currFlowInstName))
                {
                    allBinoutItems[currFlowInstName].AddRange(binoutValue.Value.ToList());
                }
                else
                {
                    allBinoutItems.Add(currFlowInstName, binoutValue.Value.ToList());
                }
            }
        }

        internal Dictionary<string, string> GetBinoutEnableDic(IEnumerable<BinOutStatusHipListRow> binOutRows)
        {
            var newBinoutStatus = new Dictionary<string, string>();
            foreach (BinOutStatusHipListRow row in binOutRows)
            {
                foreach (KeyValuePair<string, string> binoutStatus in row.BinOutEnableDictionary)
                {
                    if (newBinoutStatus.ContainsKey(binoutStatus.Key))
                    {
                        BinOutStatus status = GetBinoutStatusFormReport(newBinoutStatus[binoutStatus.Key]);
                        BinOutStatus currStatus = GetBinoutStatusFormReport(binoutStatus.Value);
                        if (!status.Equals(currStatus))
                        {
                            if (status.Equals(BinOutStatus.Unknown))
                            {
                                newBinoutStatus[binoutStatus.Key] = binoutStatus.Value;
                            }
                            else if (currStatus.Equals(BinOutStatus.Binout))
                            {
                                newBinoutStatus[binoutStatus.Key] = binoutStatus.Value;
                            }
                        }
                    }
                    else
                    {
                        newBinoutStatus.Add(binoutStatus.Key, binoutStatus.Value);
                    }
                }
            }

            return newBinoutStatus;
        }

        internal string GetBinNameByFailFalg(Dictionary<string, List<string>> bintable, HashSet<string> flags)
        {
            foreach (string flag in flags)
            {
                var currBinItems = bintable.Where(x => x.Value.Any(y => y.Equals(flag, StringComparison.CurrentCultureIgnoreCase))).ToList();
                if (currBinItems.Any())
                {
                    return currBinItems.First().Key;
                }
            }

            return "";
        }

        internal BinOutStatus GetBinoutStatusFormReport(string binoutStauts)
        {
            if (_regNoBinOut.IsMatch(binoutStauts))
            {
                return BinOutStatus.NonBinout;
            }

            if (_regBinOut.IsMatch(binoutStauts))
            {
                return BinOutStatus.Binout;
            }

            if (_regSkipBinOut.IsMatch(binoutStauts))
            {
                return BinOutStatus.SkipBinTable;
            }

            return BinOutStatus.Unknown;
        }

        internal bool UpdateBinOutValueIsSame(Dictionary<string, string> binoutDic, HashSet<string> existJobs)
        {
            for (int i = 0; i < binoutDic.Count; i++)
            {
                if (!existJobs.Contains(binoutDic.ElementAt(i).Key))
                {
                    continue;
                }

                string job1Binoutstate = binoutDic.ElementAt(i).Value;
                for (int j = i + 1; j < binoutDic.Count; j++)
                {
                    if (!existJobs.Contains(binoutDic.ElementAt(j).Key))
                    {
                        continue;
                    }

                    string job2Binoutstate = binoutDic.ElementAt(j).Value;
                    if (!job1Binoutstate.Equals(job2Binoutstate, StringComparison.CurrentCultureIgnoreCase))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        internal bool CheckJobPartIsMatch(string rowJobPart, string currectJobPart)
        {
            if (string.IsNullOrEmpty(rowJobPart))
            {
                return true;
            }

            List<string> jobparts = rowJobPart.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            foreach (string jobpart in jobparts)
            {
                if (jobpart.StartsWith("!", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (!jobpart.ToUpper().Equals(currectJobPart.ToUpper()) && !jobparts.Any(x => x.ToUpper().Equals(currectJobPart.ToUpper())))
                    {
                        return true;
                    }
                }
                else
                {
                    if (jobpart.ToUpper().Equals(currectJobPart.ToUpper()))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void InsertMeasureItemToProgram(FlowSheet sheet, int insertIdx, List<FlowRow> measureItems)
        {
            foreach (FlowRow measureItem in measureItems)
            {
                FlowRow copy = measureItem.Copy();
                copy.Job = "";
                sheet.InsertRow(insertIdx, copy);
                insertIdx++;
            }
        }

        public bool UpdateNoBinOutInfo(FlowRow item, string sheetName = "")
        {
            if (string.IsNullOrEmpty(item.FailAction))
            {
                return false;
            }

            if (sheetName.Contains(ConIvdm) && item.Parameter.StartsWith(ConIvdm) && item.Parameter.Contains(ConIvdm3X))
            {
                return false;
            }

            UpdateFailActionDic(item.Parameter, item.FailAction);
            List<string> comment1 = item.Comment1.Split('\t').ToList();
            if (comment1.Count == 0)
            {
                comment1.Add("");
            }

            if (comment1.Count == 1)
            {
                comment1.Add(item.FailAction);
            }
            else
            {
                comment1[1] = item.FailAction;
            }

            item.Comment1 = string.Join("\t", comment1);
            item.FailAction = "";
            return true;
        }

        internal string GenerateFlagNameFromTool(FlowRow row)
        {
            Match match = _regRenameFlag.Match(row.Parameter);

            if (!match.Success)
            {
                return "";
            }

            string hardIpName = row.SheetName.Replace("Flow_", "");
            string block = match.Groups["Block"].Value;
            string subBlock = match.Groups["SubBlock"].Value;
            string pattern = match.Groups["pattern"].Value;
            string voltage = match.Groups["voltage"].Value;

            return CommonGenerator.GenHardIpFlowFailFlag(hardIpName, block, subBlock, pattern, "", "", voltage, false);
        }

        private void UpdateFailActionDic(string parameter, string failFlag)
        {
            if (!_failActionDic.ContainsKey(parameter))
            {
                _failActionDic[parameter] = new HashSet<string>();
            }

            _failActionDic[parameter].Add(failFlag);
        }

        private void RefernceOrgLimits(Dictionary<string, string> orgLimitDict, Dictionary<string, string> updateLimitDict, string job)
        {
            string originalLimit = orgLimitDict[job];

            if (string.IsNullOrEmpty(originalLimit))
            {
                return;
            }

            if (!updateLimitDict.ContainsKey(job) || !string.IsNullOrEmpty(updateLimitDict[job]))
            {
                return;
            }

            updateLimitDict[job] = originalLimit;
        }

        private void ProcessUpdateTp(
            BinOutStatusHipListRow binoutItem,
            HashSet<string> currentJobs,
            HashSet<string> jobparts,
            string jobpart,
            bool isFuncItem,
            HashSet<string> funcItemJobPart,
            FlowSheet sheet,
            FlowRow flowItem,
            List<FlowRow> flowAllTestItems,
            List<FlowRow> measureItems)
        {
            HashSet<string> needUpdateJobs = GetNeedUpdateTpJob(binoutItem, currentJobs);
            if (!needUpdateJobs.Any() || !needUpdateJobs.Contains(jobpart))
            {
                return;
            }

            HashSet<string> updateSameValueJobParts = UpdateValueIsSame(binoutItem, jobpart, needUpdateJobs);
            HashSet<string> flowJobs = !string.IsNullOrEmpty(flowItem.Job) ? flowItem.Job.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x + "_").ToHashSet() : jobparts;

            if (isFuncItem)
            {
                funcItemJobPart.AddRange(updateSameValueJobParts);
                binoutItem.IsUpdated.AddRange(updateSameValueJobParts);
            }
            else
            {
                if (!funcItemJobPart.Contains(jobpart))
                {
                    return;
                }
                updateSameValueJobParts.IntersectWith(flowJobs);
                binoutItem.IsUpdated.AddRange(updateSameValueJobParts);
            }

            if (updateSameValueJobParts.Count != needUpdateJobs.Count)
            {
                bool needUpdateLimitToFunc = isFuncItem && measureItems.Any();
                int flowIndex = sheet.Rows.FindIndex(x => x.Equals(flowItem));
                if (isFuncItem)
                {
                    bool needUpdate = InsertItemToProgram(needUpdateLimitToFunc, sheet, flowIndex, flowItem, updateSameValueJobParts, binoutItem, flowAllTestItems, measureItems);
                    if (!needUpdate)
                    {
                        return;
                    }

                    var notUpdateJob = jobparts.Where(x => !updateSameValueJobParts.Contains(x)).ToHashSet();
                    flowItem.Job = string.Join(",", notUpdateJob).Replace("_", "");
                }
                else
                {
                    var notUpdateJob = flowJobs.Where(x => !binoutItem.IsUpdated.Contains(x) && funcItemJobPart.Contains(x)).ToHashSet();
                    if (notUpdateJob.Any())
                    {
                        bool needUpdate = InsertItemToProgram(needUpdateLimitToFunc, sheet, flowIndex, flowItem, updateSameValueJobParts, binoutItem, flowAllTestItems, measureItems);
                        if (!needUpdate)
                        {
                            return;
                        }

                        flowItem.Job = string.Join(",", notUpdateJob).Replace("_", "");
                    }
                    else
                    {
                        UpdateItemToProgram(ref flowItem, binoutItem, jobpart, flowAllTestItems, sheet.Name);
                    }
                }
            }
            else
            {
                UpdateItemToProgram(ref flowItem, binoutItem, jobpart, flowAllTestItems, sheet.Name);
            }
        }
    }
}
