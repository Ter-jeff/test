using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutFlowInstanceWriter
    {
        protected readonly BinCutInputData BinCutInputData;
        protected readonly List<BinCutFinalInstanceRow> BinCutFinalInstanceRow;
        protected readonly List<string> GradeSearchJobs;
        public readonly Dictionary<string, BinCutExtraPolationModeInfo> BinCutExtraPolationModes = new();
        private const string AdjustVddbinning = "Adjust_VddBinning";

        public BinCutFlowInstanceWriter(BinCutInputData binCutInputManager, List<BinCutFinalInstanceRow> binCutFinalInstanceRows, List<string> gradeSearchJobs = null)
        {
            BinCutInputData = binCutInputManager;
            BinCutFinalInstanceRow = binCutFinalInstanceRows;
            GradeSearchJobs = gradeSearchJobs;
        }

        public virtual void GetBinCutResult(Dictionary<string, List<BinCutSourceItem>> sourceDic, out BinCutPatternReport binCutPatternReport, out List<BinCutBinningItem> binCutBinningItems)
        {
            var instanceSheet = new InstanceSheet("TestInst_Vddbinning");
            var binCutFlowSheets = new List<SubFlowSheet>();

            binCutPatternReport = new BinCutPatternReport();

            IEnumerable<KeyValuePair<string, List<BinCutSourceItem>>> lvSourceRows = sourceDic.Where(x => !x.Key.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase));
            IEnumerable<KeyValuePair<string, List<BinCutSourceItem>>> hvSourceRows = sourceDic.Where(x => x.Key.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase));
            const bool isPost = false;
            var instanceRows = new InstanceRows();
            foreach (KeyValuePair<string, List<BinCutSourceItem>> item in lvSourceRows)
            {
                if (item.Value.Count == 0)
                {
                    continue;
                }

                string mode = Regex.Replace(Regex.Replace(item.Key, "^Flow_", ""), "_TD_Mbist_BV$", "");
                bool intSkipTest = BinCutInputData.BinningTables.First().IsInterpolationSkip(mode);
                var binCutInstanceAndSourceRow = new BinCutInstanceLvGenerator(item.Value, BinCutInputData, BinCutFinalInstanceRow);
                binCutInstanceAndSourceRow.GenInstanceRows(isPost, BinCutInputData);
                binCutInstanceAndSourceRow.BinCutInstanceRowMergeByJob();
                binCutInstanceAndSourceRow.BinCutInstanceNameCheck();//if the instance name is duplicated, the flag turn true
                List<BinCutRowForSort> reOrderResult = binCutInstanceAndSourceRow.ReArrangeByOrderOption(item.Key, false);
                instanceRows.AddRange(binCutInstanceAndSourceRow.GetInstanceRows(item.Key, reOrderResult, out List<BinCutPatternRow> _, isPost, intSkipTest));
                bool isCSharp = instanceRows.Any(x => x.VbtName.EndsWith("BinCutTest", StringComparison.CurrentCultureIgnoreCase));
                SubFlowSheet binCutFlowSheet = binCutInstanceAndSourceRow.GenerateFlowRows(item.Key, reOrderResult, isPost, isCSharp);
                binCutFlowSheets.Add(binCutFlowSheet);
            }

            var existHvccFlags = new List<BinCutBinningItem>();
            foreach (KeyValuePair<string, List<BinCutSourceItem>> item in hvSourceRows)
            {
                var binCutInstanceAndSourceRowHv = new BinCutInstanceHvGenerator(item.Value, BinCutInputData, BinCutFinalInstanceRow);
                binCutInstanceAndSourceRowHv.GenInstanceRows(isPost, BinCutInputData);
                binCutInstanceAndSourceRowHv.BinCutInstanceRowMergeByJob();
                binCutInstanceAndSourceRowHv.BinCutInstanceNameCheck();//if the instance name is duplicated, the flag turn true
                List<BinCutRowForSort> reOrderResult = binCutInstanceAndSourceRowHv.ReArrangeByOrderOption(item.Key, false);
                instanceRows.AddRange(binCutInstanceAndSourceRowHv.GetInstanceRows(item.Key, reOrderResult, out List<BinCutPatternRow> _, isPost, false));
                bool isCSharp = instanceRows.Any(x => x.VbtName.EndsWith("BinCutTest", StringComparison.CurrentCultureIgnoreCase));
                SubFlowSheet binCutFlowSheet = binCutInstanceAndSourceRowHv.GenerateFlowRows(item.Key, reOrderResult, isPost, isCSharp);
                binCutFlowSheets.Add(binCutFlowSheet);

                existHvccFlags.AddRange(binCutInstanceAndSourceRowHv.GetHvccFlags().FindAll(x => !existHvccFlags.Select(y => y.FlagName).Contains(x.FlagName)));
            }
            binCutBinningItems = existHvccFlags;

            #region Flow_Vddbinning
            bool isCsharp = TestProgram.VbtFunctionLib.GetFunctionByName("BinCutTest", "bincut", true).IsFound;
            SubFlowSheet vddBinningMainFlow = isCsharp ? GetVddBinningFlowCs(GetFlowSequence(sourceDic.Keys.ToList()), ref binCutFlowSheets) : GetVddBinningFlow(GetFlowSequence(sourceDic.Keys.ToList()), ref binCutFlowSheets);
            foreach (SubFlowSheet sheet in binCutFlowSheets)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirBinCut, sheet);
            }

            TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirBinCut, vddBinningMainFlow);
            #endregion

            #region Instance
            var newInstanceRows = new InstanceRows();
            newInstanceRows.AddRange(instanceRows.Where(x => !(x.FinalJobs.Count == 1 && x.FinalJobs.First().Equals("QA", StringComparison.CurrentCultureIgnoreCase))).ToList());
            newInstanceRows.AddRange(instanceRows.Where(x => x.FinalJobs.Count == 1 && x.FinalJobs.First().Equals("QA", StringComparison.CurrentCultureIgnoreCase)).ToList());
            instanceSheet.Rows = newInstanceRows;
            instanceSheet.RemoveDuplicateInstance();

            instanceSheet.Rows.Add(GenPrintConfigInstanceRow(GetHarvestFailFlag(BinCutFinalInstanceRow)));

            var searchModeList = new Dictionary<string, bool>();
            foreach (KeyValuePair<string, List<BinCutSourceItem>> flow in sourceDic)
            {
                if (flow.Key.ContainsIgnoreCase("HVCC"))
                {
                    continue;
                }
                string mode = Regex.Replace(Regex.Replace(flow.Key, "^Flow_", ""), "_TD_Mbist_BV$", "");
                bool isEvaluate = flow.Value.Any(x => x.BinValues.Any(y => y.PinContext.ContainsIgnoreCase("evaluate")));
                if (searchModeList.ContainsKey(mode))
                {
                    searchModeList[mode] = isEvaluate;
                }
                else
                {
                    searchModeList.Add(mode, isEvaluate);
                }
            }

            instanceSheet.Rows.AddRange(GenerateOtherInstanceRows(searchModeList));

            #region Header/Footer for instance
            instanceSheet.Rows.AddRange(GenHeaderFooterRows(sourceDic));
            var names = new List<string> { "Vddbinning", "Vddbinning_Judge_stored_IDS", "Read_DVFM_To_GradeVDD", "Vddbinning_Pre_Adjust_VddBinning", "Vddbinning_PrintOutVddBinning", "Vddbinning_Adjust_VddBinning", "Power_Binning" };
            foreach (string name in names)
            {
                instanceSheet.AddHeaderFooter(name);
            }

            #endregion

            #region T0TX
            if (instanceRows.Any())
            {
                var flowT0Tx = new SubFlowSheet("Flow_T0TX_PreCall");
                var levelList = instanceRows.Select(x => x.PinLevels).Distinct().ToList();
                foreach (string level in levelList)
                {
                    if (!Regex.IsMatch(level, @"levels_\w+", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    InstanceRow t0TxLv = instanceRows.First(x => x.PinLevels.Equals(level)).Copy();
                    t0TxLv.ColumnA = "DummyInstanceforT0TX";
                    t0TxLv.TestName = "Vddbinning_T0TX_" + level + "_LV";
                    t0TxLv.DcCategory = "Bincut_X_X_X";
                    t0TxLv.DcSelector = "Min";
                    instanceSheet.Rows.Add(t0TxLv);
                    AddFlowItem("", "test", t0TxLv.TestName, ref flowT0Tx);

                    InstanceRow t0TxNv = t0TxLv.Copy();
                    t0TxNv.ColumnA = "DummyInstanceforT0TX";
                    t0TxNv.TestName = "Vddbinning_T0TX_" + level + "_NV";
                    t0TxNv.DcCategory = "Bincut_X_X_X";
                    t0TxNv.DcSelector = "Typ";
                    instanceSheet.Rows.Add(t0TxNv);
                    AddFlowItem("", "test", t0TxNv.TestName, ref flowT0Tx);

                    InstanceRow t0TxHv = t0TxLv.Copy();
                    t0TxHv.ColumnA = "DummyInstanceforT0TX";
                    t0TxHv.TestName = "Vddbinning_T0TX_" + level + "_HV";
                    t0TxHv.DcCategory = "Bincut_X_X_X";
                    t0TxHv.DcSelector = "Max";
                    instanceSheet.Rows.Add(t0TxHv);
                    AddFlowItem("", "test", t0TxHv.TestName, ref flowT0Tx);
                }
                AddFlowItem("", "Return", "", ref flowT0Tx);
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirBinCut, flowT0Tx);
            }

            #endregion

            TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirBinCut, instanceSheet);
            #endregion
        }

        internal Dictionary<string, List<string>> GetFlowSequence(List<string> sourceFlowList)
        {
            var finalSourceFlow = new Dictionary<string, List<string>> { { "All", sourceFlowList } };
            return finalSourceFlow;
        }

        internal List<string> GetHarvestFailFlag(IEnumerable<BinCutFinalInstanceRow> binCutInstanceRows)
        {
            var harvestFlag = new List<string>();
            foreach (BinCutFinalInstanceRow row in binCutInstanceRows)
            {
                if (!string.IsNullOrEmpty(row.BinCutInstanceRow.SiteVar))
                {
                    if (row.BinCutInstanceRow.SiteVar.Contains("&&") || row.BinCutInstanceRow.SiteVar.Contains("||"))
                    {
                        harvestFlag.AddRange(row.BinCutInstanceRow.SiteVar.Split(new[] { "&&", "||", "!", " ", "(", ")" }, StringSplitOptions.RemoveEmptyEntries));
                    }
                    else
                    {
                        harvestFlag.AddRange(row.BinCutInstanceRow.SiteVar.Split(new[] { '!', ' ', '(', ')' }, StringSplitOptions.RemoveEmptyEntries));
                    }
                }

                if (!string.IsNullOrEmpty(row.BinCutInstanceRow.EnableAndDevice))
                {
                    string enableAndDevice = GetDeviceNameFromEnableAndDeviceCol(row.BinCutInstanceRow.EnableAndDevice);
                    if (!string.IsNullOrEmpty(enableAndDevice))
                    {
                        harvestFlag.Add(enableAndDevice);
                    }
                }

            }
            if (harvestFlag.Any())
            {
                harvestFlag = harvestFlag.Distinct().ToList();
                harvestFlag = SortHarvestFlags(harvestFlag);
            }

            return harvestFlag;
        }

        internal List<string> SortHarvestFlags(List<string> harvestFlags)
        {
            var harvList = harvestFlags.Where(x => Regex.IsMatch(x, @".HARV\d+", RegexOptions.IgnoreCase)).ToList();
            var harvListGroup = harvList.GroupBy(x => Regex.Replace(x, @"HARV\d+$", "", RegexOptions.IgnoreCase)).ToList();
            foreach (IGrouping<string, string> harvestList in harvListGroup)
            {
                var hFlags = new List<string>(harvestList.ToList());
                BubbleSortHarvestFlag(ref hFlags);
            }

            var coreList = harvestFlags.Where(x => Regex.IsMatch(x, @".CORE\d+", RegexOptions.IgnoreCase)).ToList();
            var coreListGroup = coreList.GroupBy(x => Regex.Replace(x, @"CORE\d+$", "", RegexOptions.IgnoreCase)).ToList();
            foreach (IGrouping<string, string> core in coreListGroup)
            {
                var cFlags = new List<string>(core.ToList());
                BubbleSortHarvestFlag(ref cFlags);
            }

            var totalFlags = harvestFlags.Where(flag => !Regex.IsMatch(flag, @".HARV\d+", RegexOptions.IgnoreCase) && !Regex.IsMatch(flag, @".CORE\d+", RegexOptions.IgnoreCase)).ToList();
            if (harvList.Any())
            {
                totalFlags.AddRange(harvList);
            }

            if (coreList.Any())
            {
                totalFlags.AddRange(coreList);
            }

            return totalFlags;
        }

        internal void BubbleSortHarvestFlag(ref List<string> flagList)
        {
            for (int i = 0; i < flagList.Count; i++)
            {
                for (int j = 0; j < flagList.Count - 1; j++)
                {
                    short num1 = short.Parse(Regex.Replace(flagList[j], @"\D", ""));
                    short num2 = short.Parse(Regex.Replace(flagList[j + 1], @"\D", ""));
                    if (num1 > num2)
                    {
                        (flagList[j + 1], flagList[j]) = (flagList[j], flagList[j + 1]);
                    }
                }
            }
        }

        internal string GetDeviceNameFromEnableAndDeviceCol(string enableAndDevice)
        {
            if (!string.IsNullOrEmpty(enableAndDevice))
            {
                string[] arr = enableAndDevice.Split(',');
                foreach (string data in arr)
                {
                    if (string.IsNullOrEmpty(data))
                    {
                        continue;
                    }

                    if (data.Contains("<#>"))
                    {
                        continue;
                    }

                    if (data.EndsWith("@site", StringComparison.CurrentCultureIgnoreCase))
                    {
                        return Regex.Replace(data, "@site", "", RegexOptions.IgnoreCase).Trim().TrimStart('!');
                    }
                }
            }
            return "";
        }

        protected SubFlowSheet GetVddBinningFlowCs(Dictionary<string, List<string>> subFlowSheetNameList, ref List<SubFlowSheet> binCutFlowSheets)
        {
            const string lStrSheetName = "Flow_Vddbinning";
            var flow = new SubFlowSheet(lStrSheetName, "BV");
            AddFlowItem("", OpCode.Print, "\"print: Flow_Vddbinning start\"", "", ref flow);
            AddFlowItem("", "call", "Flow_T0TX_PreCall", ref flow);
            AddFlowItem("", "test", "Vddbinning_Header_1", ref flow);
            AddFlowItem("", "Flag-Clear", BinCutConstant.VddBinningFailStopFlag, ref flow);
            AddFlowItem("", "test", "Print_BinCut_Config", ref flow);
            AddFlowItem("", "test", "Vddbinning_Judge_stored_IDS_Header_1", ref flow);
            string gradeSearchJobs = "CP1";
            if (GradeSearchJobs != null)
            {
                List<string> filterByBmsGradeSearchJobs = new List<string>();
                if (BinCutInputData.NewBinCutFlowTables != null && BinCutInputData.NewBinCutFlowTables.Any())
                {
                    var jobNames = BinCutInputData.NewBinCutFlowTables.Select(ft => ft.JobName).Distinct().ToList();
                    foreach (string gradeJob in GradeSearchJobs)
                    {
                        if (BinCutConstant.GradeJobMap.TryGetValue(gradeJob, out HashSet<string> relatedJobs))
                        {
                            string matchedJob = jobNames
                                .FirstOrDefault(name => relatedJobs.Any(rj => name.Equals(rj, StringComparison.OrdinalIgnoreCase)));

                            if (!filterByBmsGradeSearchJobs.Contains(matchedJob) && matchedJob != null)
                            {
                                filterByBmsGradeSearchJobs.Add(matchedJob);
                            }
                        }
                        else
                        {
                            filterByBmsGradeSearchJobs.Add(gradeJob);
                        }

                    }
                }
                gradeSearchJobs = string.Join("||", filterByBmsGradeSearchJobs);
                string.Join("&&", filterByBmsGradeSearchJobs.Select(job => "!" + job));
            }
            AddFlowItem("", OpCode.Print, "\"print: Judge_stored_IDS start\"", "", ref flow);
            AddFlowItem(gradeSearchJobs, "test", "Judge_stored_IDS", "", ref flow);
            AddFlowItem("", OpCode.Print, "\"print: Judge_stored_IDS end\"", "", ref flow);
            AddFlowItem("", "test", "Vddbinning_Judge_stored_IDS_Footer_1", ref flow);

            AddFlowItem("", "BinTable", "Bin_Vddbinning_IDS_fail", ref flow);
            if (subFlowSheetNameList.Count == 1)
            {
                List<string> performanceList = subFlowSheetNameList.First().Value;
                GenSubFlowItem(ref flow, performanceList);
            }
            else
            {
                foreach (KeyValuePair<string, List<string>> pair in subFlowSheetNameList)
                {
                    List<string> jobs = pair.Key.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    GenSubFlowItem(ref flow, pair.Value, jobs);
                }
            }

            AddFlowItem(gradeSearchJobs, "test", "Set_VBinResult_without_Test", ref flow);


            AddFlowItem("", "test", "Vddbinning_PrintOutVddBinning_Header_1", ref flow);
            AddFlowItem("", OpCode.Print, "\"print: Vddbinning_PrintOutVddBinning start\"", "", ref flow);
            AddFlowItem("", "test", "PrintOutVddBinning", ref flow, BinCutConstant.VddBinningFailStopFlag);
            AddFlowItem("", "BinTable", "Vddbinning_Fail_Stop", ref flow);
            AddFlowItem("", OpCode.Print, "\"print: Vddbinning_PrintOutVddBinning end\"", "", ref flow);
            AddFlowItem("", "test", "Vddbinning_PrintOutVddBinning_Footer_1", ref flow);

            AddFlowItem("", "test", "Vddbinning_Adjust_VddBinning_Header_1", ref flow);
            AddFlowItem("", OpCode.Print, "\"print: Vddbinning_Adjust_VddBinning start\"", "", ref flow);
            AddFlowItem(gradeSearchJobs, "test", AdjustVddbinning, ref flow, BinCutConstant.VddBinningFailStopFlag);
            AddFlowItem("", "BinTable", "Vddbinning_Fail_Stop", ref flow);
            AddFlowItem("", OpCode.Print, "\"print: Vddbinning_Adjust_VddBinning end\"", "", ref flow);
            AddFlowItem("", "test", "Vddbinning_Adjust_VddBinning_Footer_1", ref flow);
            AddFlowItem("", OpCode.Print, "\"print: Flow_Vddbinning end\"", "", ref flow);
            AddFlowItem("", "test", "FuseBinnedProductVoltages", ref flow);
            if (LocalSpecs.AllPowerBinningFileName != null && LocalSpecs.AllPowerBinningFileName.Count > 0)
            {
                AddFlowItem("", "test", "Power_Binning_Header_1", ref flow);
                AddFlowItem(gradeSearchJobs, "test", "Power_Binning", ref flow);
                AddFlowItem("", "BinTable", BinCutConstant.PowerBinningFail, ref flow);
                AddFlowItem("", "test", "Power_Binning_Footer_1", ref flow);
            }
            if (binCutFlowSheets.Exists(p => p.Name.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase)))
            {
                var rows = binCutFlowSheets.Where(p => p.Name.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase)).ToList();

                if (rows.Count != 1)
                {
                    SubFlowSheet sheet = binCutFlowSheets.Find(p => p.Name.Equals("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase));
                    var flowSheets = binCutFlowSheets.Where(p => p.Name.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase))
                        .Where(p => !p.Name.Equals("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase)).ToList();
                    foreach (SubFlowSheet flowSheet in flowSheets)
                    {
                        if (sheet.IsSame(flowSheet))
                        {
                            binCutFlowSheets.Remove(flowSheet);
                        }
                    }
                }
            }

            AddFlowItem("", "test", "Vddbinning_Footer_1", ref flow);
            AddFlowItem("", "Return", "", ref flow);
            return flow;
        }

        protected SubFlowSheet GetVddBinningFlow(Dictionary<string, List<string>> subFlowSheetNameList, ref List<SubFlowSheet> binCutFlowSheets)
        {
            const string lStrSheetName = "Flow_Vddbinning";
            var flow = new SubFlowSheet(lStrSheetName, "BV");
            AddFlowItem("", "call", "Flow_T0TX_PreCall", ref flow);
            AddFlowItem("", "test", "Vddbinning_Header_1", ref flow);
            AddFlowItem("", "Flag-Clear", BinCutConstant.VddBinningFailStopFlag, ref flow);
            AddFlowItem("", "test", "Print_BinCut_Config", ref flow);
            AddFlowItem("", "test", "Vddbinning_Judge_stored_IDS_Header_1", ref flow);
            string gradeSearchJobs = "CP1";
            string notGradeSearchJobs = "(CP2 || FT1 || FT2 || QA)";
            if (GradeSearchJobs != null)
            {
                gradeSearchJobs = string.Join("||", GradeSearchJobs);
                notGradeSearchJobs = string.Join("&&", GradeSearchJobs.Select(job => "!" + job));
            }

            AddFlowItem(gradeSearchJobs, "test", "Judge_stored_IDS", "F_Vddbinning_IDS_fail", ref flow);
            AddFlowItem("", "test", "Vddbinning_Judge_stored_IDS_Footer_1", ref flow);

            AddFlowItem("", "BinTable", "Bin_Vddbinning_IDS_fail", ref flow);
            AddFlowItem(notGradeSearchJobs, "test", "Read_DVFM_To_GradeVDD_Header_1", ref flow);
            AddFlowItem(notGradeSearchJobs, "test", "Read_DVFM_To_GradeVDD", ref flow);
            AddFlowItem(notGradeSearchJobs, "BinTable", "Vddbinning_Fail_Stop", ref flow);
            AddFlowItem(notGradeSearchJobs, "test", "Read_DVFM_To_GradeVDD_Footer_1", ref flow);
            if (subFlowSheetNameList.Count == 1)
            {
                List<string> performanceList = subFlowSheetNameList.First().Value;
                GenSubFlowItem(ref flow, performanceList);
            }
            else
            {
                foreach (KeyValuePair<string, List<string>> pair in subFlowSheetNameList)
                {
                    List<string> jobs = pair.Key.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    GenSubFlowItem(ref flow, pair.Value, jobs);
                }
            }

            AddFlowItem(gradeSearchJobs, "test", "Set_VBinResult_without_Test", ref flow);

            if (LocalSpecs.AllPowerBinningFileName != null && LocalSpecs.AllPowerBinningFileName.Count > 0)
            {
                AddFlowItem("", "test", "Power_Binning_Header_1", ref flow);
                AddFlowItem("", "test", "Power_Binning", ref flow);
                AddFlowItem("", "BinTable", BinCutConstant.PowerBinningFail, ref flow);
                AddFlowItem("", "test", "Power_Binning_Footer_1", ref flow);
            }
            AddFlowItem("", "test", "Vddbinning_PrintOutVddBinning_Header_1", ref flow);
            AddFlowItem("", "test", "PrintOutVddBinning", ref flow, BinCutConstant.VddBinningFailStopFlag);
            AddFlowItem("", "BinTable", "Vddbinning_Fail_Stop", ref flow);
            AddFlowItem("", "test", "Vddbinning_PrintOutVddBinning_Footer_1", ref flow);

            AddFlowItem("", "test", "Vddbinning_Adjust_VddBinning_Header_1", ref flow);
            AddFlowItem("", "test", AdjustVddbinning, ref flow, BinCutConstant.VddBinningFailStopFlag);
            AddFlowItem("", "BinTable", "Vddbinning_Fail_Stop", ref flow);
            AddFlowItem("", "test", "Vddbinning_Adjust_VddBinning_Footer_1", ref flow);

            if (binCutFlowSheets.Exists(p => p.Name.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase)))
            {
                var rows = binCutFlowSheets.Where(p => p.Name.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase)).ToList();

                if (rows.Count != 1)
                {
                    SubFlowSheet sheet = binCutFlowSheets.Find(p => p.Name.Equals("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase));
                    var flowSheets = binCutFlowSheets.Where(p => p.Name.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase))
                        .Where(p => !p.Name.Equals("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase)).ToList();
                    foreach (SubFlowSheet flowSheet in flowSheets)
                    {
                        if (sheet.IsSame(flowSheet))
                        {
                            binCutFlowSheets.Remove(flowSheet);
                        }
                    }
                }
            }

            AddFlowItem("", "test", "Vddbinning_Footer_1", ref flow);
            AddFlowItem("", "Return", "", ref flow);
            return flow;
        }

        internal string GetFinalInterpoJobs(List<string> currJobs)
        {
            var finalJobs = currJobs.Where(job => GradeSearchJobs.Exists(x => x.Equals(job, StringComparison.CurrentCultureIgnoreCase))).ToList();
            return finalJobs.Any() ? string.Join("||", finalJobs) : "";
        }


        public class PendingInterpolate
        {
            public string Interpolation;
            public string Domain;
            public string Mode;
            public string IntModeL;
            public string IntModeH;
            public string SheetName;
        }

        public string GetCurrentJob(string currentJob, List<string> jobs = null)
        {
            if (currentJob.Contains("All"))
            {
                return jobs != null && jobs.Any()
                    ? GetFinalInterpoJobs(jobs)
                    : string.Join("||", GradeSearchJobs);
            }

            string matchedJob = BinCutConstant.GradeJobMap
                .FirstOrDefault(kvp => kvp.Value
                    .Contains(currentJob, StringComparer.OrdinalIgnoreCase)).Key;

            string target = matchedJob ?? currentJob;
            bool exists = GradeSearchJobs.Any(i => i.ContainsIgnoreCase(target));
            return exists ? currentJob : string.Empty;
        }

        public (bool, List<string>) HandleBinTablesJobs(
            int index,
            BinningTables tables,
            string oriMode,
            string sheetName,
            List<PendingInterpolate> pendingInterpolate,
            List<string> performanceList
        )
        {
            List<string> allCurrentJobs = new List<string>();
            bool isInterpolate = false;
            bool isMultiRails = false;
            string key = null;
            var interpoKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int evaluateBinCount = 0;

            List<string> pModeList = GetMultiRailPerformanceModes(oriMode, BinCutInputData.BinCutFlowTables.First().Rows);
            evaluateBinCount = pModeList.Count;
            if (evaluateBinCount <= 1)
            {
                pModeList.Clear();
                pModeList.Add(oriMode);
            }
            else
            {
                isMultiRails = true;
            }

            foreach (BinningTable table in tables)
            {
                foreach (string mode in pModeList)
                {
                    if (!table.IsInterpolationSkip(mode, out string domain, out string intModeL, out string intModeH, out string _))
                    {
                        continue;
                    }
                    isInterpolate = true;
                    string currentJob = GetCurrentJob(table.Job);
                    if (string.IsNullOrEmpty(currentJob))
                    {
                        continue;
                    }

                    if (!allCurrentJobs.Any(j => j.Split(new[] { "||" }, StringSplitOptions.None).Contains(currentJob)))
                    {
                        allCurrentJobs.Add(currentJob);
                    }

                    if (BinCutExtraPolationModes != null && BinCutExtraPolationModes.TryGetValue(mode, out BinCutExtraPolationModeInfo modeInfo))
                    {
                        BinCutExtraPolationModes[mode].IsExtraPolation = true;
                        BinCutExtraPolationModes[mode].IsOnlyExtraMode = false;
                        key = $"EPL_VDD_{domain}_{mode}_BV";
                    }
                    else
                    {
                        key = $"IPL_VDD_{domain}_{mode}_BV";
                    }

                    if (interpoKeys.Add(key))
                    {
                        pendingInterpolate.Add(new PendingInterpolate
                        {
                            Interpolation = "",
                            Domain = domain,
                            Mode = mode,
                            IntModeL = intModeL,
                            IntModeH = intModeH,
                            SheetName = sheetName
                        });
                    }
                    string modeL = "Flow_" + intModeL + "_TD_Mbist_BV";
                    string modeH = "Flow_" + intModeH + "_TD_Mbist_BV";
                    List<string> flows = performanceList.GetRange(0, index + 1);
                    if ((!flows.Any(x => x.Equals(modeL, StringComparison.CurrentCultureIgnoreCase)) ||
                        !flows.Any(x => x.Equals(modeH, StringComparison.CurrentCultureIgnoreCase))) && !isMultiRails)
                    {
                        if (!flows.Any(x => x.Equals(modeL, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            ErrorReportManager.AddError(BinCutErrorType.E_Missing_02, modeL, index, 16, [key, modeL]);
                        }
                        if (!flows.Any(x => x.Equals(modeH, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            ErrorReportManager.AddError(BinCutErrorType.E_Missing_02, modeH, index, 16, [key, modeH]);
                        }
                        Response.Report("Please check flow seq. for Interpolation !!!", EnumMessageLevel.Error, 0);
                    }
                }
            }
            return (isInterpolate, allCurrentJobs);
        }

        public (bool, List<string>) HandleExtraBinTablesJobs(IdsDistributionTable startEQNtables, string mode)
        {
            List<string> allCurrentJobs = new List<string>();
            bool isExtrapolate = false;

            var jobs = startEQNtables.AllIdsPowers
                .Where(p =>
                    !string.IsNullOrWhiteSpace(p.PowerName) &&
                    p.PowerName.StartsWith(mode, StringComparison.OrdinalIgnoreCase) &&
                    p.IdsInfos != null &&
                    p.IdsInfos.Any(info =>
                        info.Polation != null &&
                        info.Polation.Any(pol =>
                            !string.IsNullOrWhiteSpace(pol) &&
                            pol.Trim().StartsWith("extra", StringComparison.OrdinalIgnoreCase))
                    )
                )
                .Select(p => p.Job)
                .Where(j => !string.IsNullOrWhiteSpace(j))
                .Distinct()
                .ToList();

            foreach (string job in jobs)
            {
                string currentJob = GetCurrentJob(job);
                if (string.IsNullOrEmpty(currentJob))
                {
                    continue;
                }

                if (!allCurrentJobs.Any(j => j.Split(new[] { "||" }, StringSplitOptions.None).Contains(currentJob)))
                {
                    allCurrentJobs.Add(currentJob);
                }
            }
            if (allCurrentJobs.Count > 0)
            {
                isExtrapolate = true;
                if (!BinCutExtraPolationModes.ContainsKey(mode))
                {
                    BinCutExtraPolationModes[mode] = new BinCutExtraPolationModeInfo();
                    BinCutExtraPolationModes[mode].IsOnlyExtraMode = true;
                    BinCutExtraPolationModes[mode].IsExtraPolation = false;
                }
            }
            return (isExtrapolate, allCurrentJobs);
        }

        internal static (string, string) ProcessJobsResult(List<string> allCurrentJobs)
        {
            string resultInterpoJobs = "";
            string resultRegularJobs = "";
            if (allCurrentJobs.Count == 1)
            {
                resultRegularJobs = "!" + allCurrentJobs[0];
                resultInterpoJobs = allCurrentJobs[0];
            }
            else if (allCurrentJobs.Count > 1)
            {
                resultRegularJobs = "!(" + string.Join("||", allCurrentJobs) + ")";
                resultInterpoJobs = string.Join("||", allCurrentJobs);
            }
            return (resultRegularJobs, resultInterpoJobs);
        }

        /// <summary>
        /// Add pending interpolations to flow item and clear pending interpolations.
        /// </summary>
        /// <param name="pendingInterpolate">Pending interpolations that need to be processed</param>
        /// <param name="flow"></param>
        /// <param name="isCsharp">Identify target is C# or not</param>
        private static void ProcessPendingVddInterpolations(
            List<PendingInterpolate> pendingInterpolate,
            ref SubFlowSheet flow,
            bool isCsharp,
            Dictionary<string, BinCutExtraPolationModeInfo> binCutExtraPolationModeList
        )
        {
            foreach (PendingInterpolate p in pendingInterpolate)
            {
                string pParam = isCsharp
                    ? BinCutConstant.VddBinningInterpolationFailCs
                    : BinCutConstant.VddBinningInterpolationFail;
                AddFlowItem(p.Interpolation, "print", $"\"****Start of calculation for Vx_{p.Domain}_{p.Mode}****\"", ref flow);
                if (binCutExtraPolationModeList != null && binCutExtraPolationModeList.Any(x => x.Key.Equals(p.Mode, StringComparison.OrdinalIgnoreCase)))
                {
                    AddFlowItem(p.Interpolation, "Test", $"EPL_VDD_{p.Domain}_{p.Mode}_BV", ref flow, "", p.IntModeL + "," + p.IntModeH);
                }
                else
                {
                    AddFlowItem(p.Interpolation, "Test", $"IPL_VDD_{p.Domain}_{p.Mode}_BV", ref flow, "", p.IntModeL + "," + p.IntModeH);
                }
                AddFlowItem(p.Interpolation, OpCode.BinTable, pParam, ref flow);
                AddFlowItem(p.Interpolation, "print", $"\"****End of calculation for Vx_{p.Domain}_{p.Mode}****\"", ref flow);
            }
            pendingInterpolate.Clear();
        }

        private void GenSubFlowItem(ref SubFlowSheet flow, List<string> performanceList, List<string> jobs = null)
        {
            bool isCsharp = TestProgram.VbtFunctionLib.GetFunctionByName("BinCutTest", "bincut", true).IsFound;
            string currJobs = jobs != null ? string.Join("||", jobs) : "";
            var pendingInterpolate = new List<PendingInterpolate>();
            string currentGroup = null;

            for (int index = 0; index < performanceList.Count; index++)
            {
                bool startWithFlow = performanceList
                    .ElementAt(index)
                    .StartsWith("Flow_", StringComparison.CurrentCultureIgnoreCase);
                string sheetName = startWithFlow
                    ? performanceList.ElementAt(index)
                    : "Flow_" + performanceList.ElementAt(index) + "_TD_Mbist_BV";
                if (sheetName.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }
                string mode = Regex.Replace(Regex.Replace(sheetName, "^Flow_", ""), "_TD_Mbist_BV$", "");
                Match m = Regex.Match(mode, @"^([A-Za-z]+)\d+");
                string group = m.Success ? m.Groups[1].Value : mode;

                (bool isExtrapolate, List<string> allExtraCurrentJobs) = HandleExtraBinTablesJobs(
                   BinCutInputData.IdsdistributionTable,
                   mode
                );

                if (currentGroup != null && !string.Equals(group, currentGroup, StringComparison.OrdinalIgnoreCase))
                {
                    ProcessPendingVddInterpolations(pendingInterpolate, ref flow, isCsharp, BinCutExtraPolationModes);
                }
                currentGroup = group;

                (bool isInterpolate, List<string> allCurrentJobs) = HandleBinTablesJobs(
                    index,
                    BinCutInputData.BinningTables,
                    mode, sheetName,
                    pendingInterpolate,
                    performanceList
                );

                if (!isInterpolate && isExtrapolate)
                {
                    (string resultExtrapolateJobs, string resultExtrapoJobs) = ProcessJobsResult(allExtraCurrentJobs);
                    foreach (PendingInterpolate p in pendingInterpolate)
                    {
                        if (string.IsNullOrEmpty(p.Interpolation))
                        {
                            p.Interpolation = resultExtrapoJobs;
                        }
                    }
                    AddFlowItem(resultExtrapolateJobs, OpCode.Call, sheetName, ref flow);
                    continue;
                }
                else if (!isInterpolate)
                {
                    AddFlowItem(currJobs, OpCode.Call, sheetName, ref flow);
                    continue;
                }

                (string resultRegularJobs, string resultInterpoJobs) = ProcessJobsResult(allCurrentJobs);

                foreach (PendingInterpolate p in pendingInterpolate)
                {
                    if (string.IsNullOrEmpty(p.Interpolation))
                    {
                        p.Interpolation = resultInterpoJobs;
                    }
                }
                AddFlowItem(resultRegularJobs, OpCode.Call, sheetName, ref flow);

            }
            ProcessPendingVddInterpolations(pendingInterpolate, ref flow, isCsharp, BinCutExtraPolationModes);
        }

        public InstanceRows GenHeaderFooterRows(Dictionary<string, List<BinCutSourceItem>> sourceRowDic)
        {
            var instanceRows = new InstanceRows();
            foreach (KeyValuePair<string, List<BinCutSourceItem>> item in sourceRowDic)
            {
                string name = item.Key.Replace("Flow_", "");
                instanceRows.AddHeaderFooter(name);
            }
            return instanceRows;
        }

        public static void AddFlowItem(string pEnable, string pOpcode, string pParameter, ref SubFlowSheet pFlowSheet, string failAction = "", string columnA = "")
        {
            var lRow = new FlowRow { Enable = pEnable, Opcode = pOpcode };
            if (pParameter.StartsWith("Vddbinning_T0TX_"))
            {
                lRow.Env = "T0TX_Use";
            }

            lRow.Parameter = pParameter;
            if (!string.IsNullOrEmpty(failAction))
            {
                lRow.FailAction = failAction;
            }

            lRow.ColumnA = columnA;
            pFlowSheet.AddRow(lRow);
        }

        internal void AddFlowItem(string pEnable, string pOpcode, string pParameter, string pFlag, ref SubFlowSheet pFlowSheet)
        {
            var lRow = new FlowRow { Enable = pEnable, Opcode = pOpcode, Parameter = pParameter, FailAction = pFlag };
            pFlowSheet.AddRow(lRow);
        }

        public List<InstanceRow> GenerateOtherInstanceRows(Dictionary<string, bool> flowNameList)
        {
            var instanceRows = new List<InstanceRow> { GenCheck_IDS() };

            if (!TestProgram.VbtFunctionLib.GetFunctionByName("BinCutTest", "bincut", true).IsFound)
            {
                AddInstanceItem(new StringBuilder().Append("Read_DVFM_To_GradeVDD").ToString(), "VBT", "Read_DVFM_To_GradeVDD", ref instanceRows);
            }

            instanceRows.Add(GenSet_VBinResult_without_Test(flowNameList, BinCutExtraPolationModes));

            if (LocalSpecs.AllPowerBinningFileName != null && LocalSpecs.AllPowerBinningFileName.Count > 0)
            {
                instanceRows.Add(GenPower_Binning_Calculation());
            }

            instanceRows.Add(GenPrintOutVddBinning());

            instanceRows.Add(GenAdjustVddBinningInstance(AdjustVddbinning));

            InstanceRow fuseBinnedProductVoltagesInstance = GenFuseBinnedProductVoltagesInstanceRow();
            if (fuseBinnedProductVoltagesInstance != null)
            {
                instanceRows.Add(fuseBinnedProductVoltagesInstance);
            }

            foreach (string sheetName in flowNameList.Keys)
            {
                if (sheetName.StartsWith("Flow_VddBinning_HVCC", StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                string oriMode = Regex.Replace(Regex.Replace(sheetName, "^Flow_", ""), "_TD_Mbist_BV$", "");
                var interpoKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                List<string> pModeList = GetMultiRailPerformanceModes(oriMode, BinCutInputData.BinCutFlowTables.First().Rows);

                if (pModeList.Count <= 1)
                {
                    pModeList = new List<string> { oriMode };
                }

                ProcessInterpolation(pModeList, interpoKeys, ref instanceRows);
            }

            return instanceRows;
        }
        private void ProcessInterpolation(IEnumerable<string> modes, HashSet<string> interpoKeys, ref List<InstanceRow> instanceRows)
        {
            foreach (BinningTable table in BinCutInputData.BinningTables)
            {
                foreach (string mode in modes)
                {
                    if (!table.IsInterpolationSkip(mode, out string domain, out _, out _, out _))
                    {
                        continue;
                    }

                    string power = $"VDD_{domain}_{mode}";
                    AddInterpolationInstance(mode, power, interpoKeys, ref instanceRows);
                }
            }
        }

        private void AddInterpolationInstance(string mode, string power, HashSet<string> interpoKeys, ref List<InstanceRow> instanceRows)
        {
            if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                var args = new List<string> { power };
                AddInstanceItem($"IPL_{power}_BV", "VBT", "ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2", ref instanceRows, "powerDomain", args);
                return;
            }

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("BinCutInterpolation", "bincut", true);
            string key;

            if (BinCutExtraPolationModes != null && BinCutExtraPolationModes.TryGetValue(mode, out BinCutExtraPolationModeInfo modeInfo) && modeInfo.IsExtraPolation)
            {
                key = $"EPL_{power}_BV";
            }
            else
            {
                key = $"IPL_{power}_BV";
            }

            if (!interpoKeys.Add(key))
            {
                return;
            }

            if (function.Type == ".NET")
            {
                var args = new List<string> { mode };
                AddInstanceItem(key, ".NET", function.FullFunctionName, ref instanceRows, "performanceMode", args);
            }
            else
            {
                var args = new List<string> { power };
                AddInstanceItem(key, "VBT", "ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2", ref instanceRows, "powerDomain", args);
            }
        }
        internal virtual InstanceRow GenPower_Binning_Calculation()
        {
            var instanceRow = new InstanceRow { TestName = "Power_Binning", VbtType = "VBT", VbtName = "Power_Binning_Calculation" };

            return instanceRow;
        }

        internal virtual InstanceRow GenCheck_IDS()
        {
            var instanceRow = new InstanceRow { TestName = "Judge_stored_IDS", VbtType = "VBT", VbtName = "check_IDS" };

            return instanceRow;
        }

        internal virtual InstanceRow GenPrintOutVddBinning()
        {
            var instanceRow = new InstanceRow { TestName = "PrintOutVddBinning", VbtType = "VBT", VbtName = "PrintOut_VDD_Bin" };

            return instanceRow;
        }

        internal virtual InstanceRow GenSet_VBinResult_without_Test(Dictionary<string, bool> flowNameList, Dictionary<string, BinCutExtraPolationModeInfo> binCutExtraPolationModeList = null)
        {
            var instanceRow = new InstanceRow
            {
                TestName = "Set_VBinResult_without_Test",
                VbtType = "VBT",
                VbtName = "Set_VBinResult_without_Test"
            };

            return instanceRow;
        }

#nullable enable
        internal virtual InstanceRow? GenFuseBinnedProductVoltagesInstanceRow()
        {
            return null;
        }
#nullable restore

        internal virtual InstanceRow GenPrintConfigInstanceRow(IEnumerable<string> harvestFlags)
        {
            var instanceRow = new InstanceRow { TestName = "Print_BinCut_config", VbtType = "VBT" };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Print_BinCut_config", "bincut");
            function.SetParamValue("str_flag_Group", string.Join(",", harvestFlags));
            instanceRow.VbtName = function.FunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;

            return instanceRow;
        }

        internal virtual InstanceRow GenAdjustVddBinningInstance(string adjustVddbinning)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(adjustVddbinning, "Bincut");
            VddBinDefComment vddBinComment = BinCutInputData.BinningTables.First().GetAdjustPowerList();
            if (vddBinComment != null)
            {
                if (vddBinComment.IsMax && vddBinComment.MaxStr != null)
                {
                    function.SetParamValue("Adjust_Max_Enable", "TRUE");
                    function.SetParamValue("Adjust_Power_Max_list", vddBinComment.MaxStr);
                    var instanceRow = new InstanceRow
                    {
                        TestName = adjustVddbinning,
                        VbtType = "VBT",
                        VbtName = function.FunctionName,
                        ArgList = function.Parameters,
                        Args = function.ArgList
                    };
                    return instanceRow;
                }

                if (vddBinComment.IsMin && vddBinComment.MinStr != null)
                {
                    function.SetParamValue("Adjust_Min_Enable", "TRUE");
                    function.SetParamValue("Adjust_Power_Min_list", vddBinComment.MinStr);
                    var instanceRow = new InstanceRow
                    {
                        TestName = adjustVddbinning,
                        VbtType = "VBT",
                        VbtName = function.FunctionName,
                        ArgList = function.Parameters,
                        Args = function.ArgList
                    };
                    return instanceRow;
                }
                else
                {
                    var instanceRow = new InstanceRow { TestName = adjustVddbinning, VbtType = "VBT", VbtName = adjustVddbinning };
                    if (!string.IsNullOrEmpty(function.Parameters))
                    {
                        instanceRow.ArgList = function.Parameters;
                    }

                    return instanceRow;
                }
            }

            {
                var instanceRow = new InstanceRow { TestName = adjustVddbinning, VbtType = "VBT", VbtName = adjustVddbinning };
                if (!string.IsNullOrEmpty(function.Parameters))
                {
                    instanceRow.ArgList = function.Parameters;
                }

                return instanceRow;
            }
        }

        internal void AddInstanceItem(string pTestName, string pType, string pName, ref List<InstanceRow> pInstanceRows, string argList = "", List<string> args = null)
        {
            var instanceRow = new InstanceRow { TestName = pTestName, VbtType = pType, VbtName = pName };
            if (!string.IsNullOrEmpty(argList))
            {
                instanceRow.ArgList = argList;
            }

            if (args != null)
            {
                instanceRow.Args = args;
            }

            pInstanceRows.Add(instanceRow);
        }

        internal List<string> GetMultiRailPerformanceModes(string oriMode, IEnumerable<BinCutFlowSheetRow> rows = null)
        {
            return rows
                .Where(r => r.TableType.Equals(EnumBinCutTableType.Lv)
                            && r.PerformanceMode.Equals(oriMode.Trim(), StringComparison.OrdinalIgnoreCase))
                .SelectMany(r => r.PinInfos)
                .Where(p => p.PinContext.IndexOf("evaluate bin",
                                    StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(p =>
                {
                    int index = p.PinContext.IndexOf(
                        "evaluate bin",
                        StringComparison.OrdinalIgnoreCase);

                    return p.PinContext.Substring(0, index).Trim();
                })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }
}
