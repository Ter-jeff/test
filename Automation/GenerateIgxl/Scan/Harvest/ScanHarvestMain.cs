using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.PostAction.GenMainFlow.Base;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.BinNumber;
using TestPlanLib.Efuse.Input;
using TestPlanLib.Harvest;
using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Scan.Harvest
{
    internal class ScanHarvestMain
    {
        private const string FlowHarvestingDecisionBlock = "Flow_Harv_Dec";
        public const string HarvestEfuseRead = "Harvest_eFuse_Read";
        public const string HarvestPostCheck = "Harvest_Post_Check";
        public const string CopyFromFuseToFlag = "Copy_From_Fuse_To_Flag";
        public const string CheckFromFuseToFlag = "Check_From_Fuse_To_Flag";
        public const string CopyFromFlagOrValueToFuse = "Copy_From_Flag_Or_Value_To_Fuse";
        public const string HarvestSummary = "Harvest_Summary";
        public const string CSharpHarvestSummary = "HarvestSummary";
        private const string HarvestSummaryAfterHarvestingDecision = "Harvest_Summary_After_Harvesting_Decision";
        private const string HarvestSummaryBeforeHarvestingDecision = "Harvest_Summary_Before_Harvesting_Decision";
        private const string HarvestSummaryAfterInit = "Harvest_Summary_After_Init";
        private readonly HarvestingTruthTableSheet _truthTable;
        private readonly Dictionary<string, string> _coreOverFailingFlows;
        private const string FHarvestOtherCombination = "F_Harvest_Other_Combination";
        private const string FHarvestReadBackMisMatch = "F_Harvest_ReadBack_MisMatch";
        private const string FHarvestPostCheckFailed = "F_Harvest_PostCheck_Fail";
        private const string BinHarvestPostCheckFail = "Bin_Harvest_PostCheck_Fail";
        private const string BinHarvestReadBackMisMatch = "Bin_other_Harvest_ReadBack_MisMatch";
        private const string BinHarvestOtherCombination = "Bin_other_Harvest_Unexpected_Combination_Failure";
        public const string HarvestSummaryAfterPrefix = "Harvest_Summary_After";
        public const string HarvestSummaryAfterTableMainInitEnableWd = HarvestSummaryAfterPrefix + "_Table_Main_Init_EnableWd";
        public const string HarvestSummaryCoreOverFailing = "Harvest_Summary_Core_Over_Failing";
        public const string RunBinOutCalcScanMbist = "Run_BinOut_Calc_Scan_Mbist";
        public Dictionary<string, string> ExistHarvBinTable = new Dictionary<string, string>();
        private const string Regex1 = "(?<flag>.+)";
        private const string Regex2 = @"(?<cores>\[.+\])";
        private static readonly Regex _regex2 = new Regex(Regex1, RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex4 = new Regex(Regex2, RegexOptions.IgnoreCase | RegexOptions.Compiled);
        public bool HasHarvestPostCheckFlow
        {
            get
            {
                MainFlowBase firstRow = TestPlanStatic.MainFlowSheet.Rows.FirstOrDefault();
                List<FlowSequenceNew> sequences = firstRow?.SequencesNew;
                return sequences != null && sequences.Any(s => s.SheetName.Equals("Flow_HarvestPostCheck", StringComparison.OrdinalIgnoreCase));
            }
        }

        public ScanHarvestMain(HarvestingTruthTableSheet truthTable)
        {
            _truthTable = truthTable;
            _coreOverFailingFlows = TestPlanStatic.MainFlowSheet.CoreOverFailingFlows;
        }

        public void GenHarvesting(List<SubFlowSheet> subFlowSheets, List<InstanceRow> instanceRows)
        {
            var decisionStatement = new List<string> { "RD", "OB", "BC" };
            if (_truthTable != null)
            {
                if (_truthTable.IndexDomain != -1)
                {
                    subFlowSheets.AddRange(GenHarvestFlows()); //Flow_Harvesting_Decision_Block_OB
                }

                instanceRows.AddRange(GenHarvestInstance());
                foreach (string statement in decisionStatement)
                {
                    if (!instanceRows.Exists(x => x.TestName.Equals(HarvestSummaryAfterHarvestingDecision + "_" + statement)))
                    {
                        GenHarvestSummaryInstance(instanceRows, HarvestSummaryAfterHarvestingDecision + "_" + statement);
                    }

                    if (!instanceRows.Exists(x => x.TestName.Equals(HarvestSummaryBeforeHarvestingDecision + "_" + statement)))
                    {
                        GenHarvestSummaryInstance(instanceRows, HarvestSummaryBeforeHarvestingDecision + "_" + statement);
                    }
                }
                if (!instanceRows.Exists(x => x.TestName.Equals(HarvestSummaryAfterTableMainInitEnableWd)))
                {
                    GenHarvestSummaryInstance(instanceRows, HarvestSummaryAfterTableMainInitEnableWd);
                }

                if (!instanceRows.Exists(x => x.TestName.Equals(HarvestSummaryAfterInit)))
                {
                    GenHarvestSummaryInstance(instanceRows, HarvestSummaryAfterInit);
                }

                foreach (string block in _coreOverFailingFlows.Values)
                {
                    GenHarvestSummaryInstance(instanceRows, string.IsNullOrEmpty(block) ? HarvestSummaryCoreOverFailing : $"{HarvestSummaryCoreOverFailing}_{block}");
                }

                if (_truthTable.SetBinFuseBit != -1 && !subFlowSheets.Exists(x => x.Name.Equals("Flow_SetBinFuse")))
                {
                    subFlowSheets.Add(CreateBinFuseFlowSheet());
                }

                if (!instanceRows.Exists(x => x.TestName.Equals(RunBinOutCalcScanMbist)))
                {
                    GenRunBinOutCalcScanMbistInstance(instanceRows, RunBinOutCalcScanMbist);
                }
            }
        }

        public List<BinTableRow> GenBinTableRows()
        {
            var totalBinTableRows = new List<BinTableRow>();
            List<BinTableRow> passBinTableRows = GenPassBinTableRows();

            SubFlowSheet mainEndFlowSheet = TestProgram.IgxlWorkBk.GetFlowSheet(IgxlWorkBook.FlowTableMainEndFlow, FolderStructure.DirMain);

            mainEndFlowSheet.InsertBinTableBeforeReturnRow(passBinTableRows);
            mainEndFlowSheet.InsertBeforeReturnRow(new List<FlowRow> { new FlowRow() { Opcode = OpCode.Stop } });
            TestProgram.IgxlWorkBk.SetFlowSheet(IgxlWorkBook.FlowTableMainEndFlow, mainEndFlowSheet);

            totalBinTableRows.AddRange(passBinTableRows);

            if (_coreOverFailingFlows.Any())
            {
                List<BinTableRow> failBinTableRows = GenFailBinTableRows();
                totalBinTableRows.AddRange(failBinTableRows);
                foreach (string block in _coreOverFailingFlows.Values)
                {
                    GenOverFailingFlowSheet(block);
                }
            }

            List<string> flags = _truthTable.Rows.First().GetSumFlags();
            foreach (BinTableRow binRow in totalBinTableRows)
            {
                flags.AddRange(binRow.ItemList.Split(',').ToList());
            }

            flags.AddRange(_truthTable.GetAllHarvFlagList());
            flags.AddRange(_truthTable.HarvFailCoreSumAllFlags());
            flags = flags.Distinct().ToList();
            flags.Sort();
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(flags, "Harvest", FolderStructure.DirMain);
            return totalBinTableRows;
        }

        public BinTableRow GenHarvestOtherBinTableRow()
        {
            var binTableRow = new BinTableRow
            {
                Name = BinHarvestOtherCombination,
                Op = "AND",
                ItemList = FHarvestOtherCombination,
                Items = new List<string> { "T" }
            };
            BinNumResult binInfo = BinNumberSingleton.Instance.GetBinInfo("Harvest", "", "", binTableRow);
            binTableRow.Sort = binInfo.SoftBin.ToString("G15");
            binTableRow.Bin = binInfo.BinNumInfo.HardBin.ToString("G15");
            binTableRow.Result = binInfo.BinNumInfo.Status;
            return binTableRow;
        }

        public BinTableRow GenHarvestReadBackBinTableRow()
        {
            var binTableRow = new BinTableRow
            {
                Name = BinHarvestReadBackMisMatch,
                Op = "AND",
                ItemList = FHarvestReadBackMisMatch,
                Items = new List<string> { "T" }
            };
            BinNumResult binInfo = BinNumberSingleton.Instance.GetBinInfo("Harvest", "", "", binTableRow);
            binTableRow.Sort = binInfo.SoftBin.ToString("G15");
            binTableRow.Bin = binInfo.BinNumInfo.HardBin.ToString("G15");
            binTableRow.Result = binInfo.BinNumInfo.Status;
            return binTableRow;
        }

        public BinTableRow GenHarvestPostCheckBinTable()
        {
            var binTableRow = new BinTableRow
            {
                Name = BinHarvestPostCheckFail,
                Op = "AND",
                ItemList = FHarvestPostCheckFailed,
                Items = new List<string> { "T" }
            };
            BinNumResult binInfo = BinNumberSingleton.Instance.GetBinInfo("Harvest", "", "", binTableRow);
            binTableRow.Sort = binInfo.SoftBin.ToString("G15");
            binTableRow.Bin = binInfo.BinNumInfo.HardBin.ToString("G15");
            binTableRow.Result = binInfo.BinNumInfo.Status;
            return binTableRow;
        }

        private string CheckEfuseField(HarvestingTruthTableRow row, string type)
        {
            ExcelWorksheet efuseBitDefTable = EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_BitDef_Table"];
            bool isMismatch = false;
            IEnumerable<string> truthTableField = type == "write" ? row.Fusings.Select(x => x.GetAddress()) : row.ReadFuse.Select(x => x.GetAddress());
            if (efuseBitDefTable != null)
            {
                Response.Report($"Reading {efuseBitDefTable.Name} ...", EnumMessageLevel.General, 20);
                List<BitDefTable> efuseBitDefTables = TestPlanStatic.BitDefTables;
                int typeIdx = efuseBitDefTables.First().DefaultOrRealIdx;
                int fieldIdx = efuseBitDefTables.First().BankEfuseBitDefIdx;
                var efuseFields = efuseBitDefTables.SelectMany(x => x.Rows).Where(x => x.RowData[typeIdx].Equals("Real", StringComparison.OrdinalIgnoreCase)).Select(x => x.RowData[fieldIdx]).ToList();
                foreach (string field in truthTableField)
                {
                    if (!efuseFields.Exists(x => x.Equals(field.Split('[').FirstOrDefault(), StringComparison.OrdinalIgnoreCase)))
                    {
                        int columnIdx;
                        if (type == "write")
                        {
                            columnIdx = _truthTable.IndexFusing + row.Fusings.FindIndex(x => x.GetAddress().Equals(field));
                        }
                        else
                        {
                            columnIdx = _truthTable.IndexReadFuse + row.ReadFuse.FindIndex(x => x.GetAddress().Equals(field));
                        }
                        _truthTable.AddError(HarvestErrorType.E_MissingField_01, _truthTable.SheetName, _truthTable.StartRowNum + 2, columnIdx, [field]);
                        isMismatch = true;
                    }
                }
            }
            else
            {
                return "";
            }
            return isMismatch ? "Efuse Fields mismatch" : "";

        }

        private List<InstanceRow> GenHarvestInstance()
        {
            var instances = new List<InstanceRow>();
            if (TestPlanStatic.HarvestingFuseWriteSheet == null)
            {
                foreach (HarvestingTruthTableRow row in _truthTable.Rows)
                {
                    if (!row.Fusings.Any())
                    {
                        continue;
                    }

                    var instanceRow = new InstanceRow
                    {
                        TestName = row.GetConditionInstacne(_truthTable.Job),
                        VbtType = "VBT",
                        VbtName = "Harvest_eFuse_Write"
                    };
                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName(instanceRow.VbtName, "scan");
                    function.SetParamValue("FuseBlockName", string.Join(";", row.Fusings.Select(x => x.Field)));
                    function.SetParamValue("FuseCategoryName", string.Join(";", row.Fusings.Select(x => x.GetAddress())));
                    function.SetParamValue("FailFlagOrValue", string.Join(";", row.Fusings.Select(x => x.Value)));
                    instanceRow.ArgList = function.Parameters;
                    instanceRow.Args = function.ArgList;
                    instanceRow.ColumnA = CheckEfuseField(row, "write");
                    instances.Add(instanceRow);
                }
            }
            else
            {
                if (_truthTable.Rows.First().Fusings.Any())
                {
                    var instanceRow = new InstanceRow();
                    List<Fusing> fusing = _truthTable.Rows.First().Fusings;
                    instanceRow.TestName = "Harvest_eFuse_Write_" + _truthTable.Job;

                    Function function = TestProgram.VbtFunctionLib.GetFunctionByName(CopyFromFlagOrValueToFuse, "scan");
                    if (function.Type == ".NET")
                    {
                        SetWriteFuseArgumentCs(ref function, fusing);
                    }
                    else
                    {
                        SetWriteFuseArgumentVbt(ref function, fusing);
                    }

                    instanceRow.VbtType = function.Type;
                    instanceRow.VbtName = function.FullFunctionName;
                    instanceRow.ArgList = function.Parameters;
                    instanceRow.Args = function.ArgList;
                    instanceRow.ColumnA = CheckEfuseField(_truthTable.Rows.FirstOrDefault(), "write");
                    instances.Add(instanceRow);
                }
            }
            string harvestEFuseReadName = HarvestEfuseRead;


            if (!string.IsNullOrEmpty(_truthTable.Job))
            {
                harvestEFuseReadName += "_" + _truthTable.Job;
            }

            if (_truthTable.Rows.First().ReadFuse.Any())
            {
                //Function Harvest_eFuse_Read(ByVal PatternName As String, ByVal FuseBlockName As String,
                //ByVal FuseCategoryName As String, ByVal FailFlagOrValue As String, ByVal HarvSrcKey As String) As Long
                var harvestEFuseReadRow = new InstanceRow { TestName = harvestEFuseReadName };
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(CopyFromFuseToFlag, "scan");
                if (function.Type == ".NET")
                {
                    SetReadFuseArgumentCs(ref function);
                }
                else
                {
                    SetReadFuseArgumentVbt(ref function);
                }

                harvestEFuseReadRow.VbtType = function.Type;
                harvestEFuseReadRow.VbtName = function.FullFunctionName;
                harvestEFuseReadRow.ArgList = function.Parameters;
                harvestEFuseReadRow.Args = function.ArgList;
                harvestEFuseReadRow.ColumnA = CheckEfuseField(_truthTable.Rows.First(), "read");
                instances.Add(harvestEFuseReadRow);
            }
            if (_truthTable.Rows.First().ReadFuse.Any() && HasHarvestPostCheckFlow)
            {
                var harvestPostCheckRow = new InstanceRow { TestName = harvestEFuseReadName.Replace(HarvestEfuseRead, HarvestPostCheck) };
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(CheckFromFuseToFlag, "scan");
                if (function.Type == ".NET")
                {
                    SetReadFuseArgumentCs(ref function);
                    harvestPostCheckRow.VbtType = function.Type;
                    harvestPostCheckRow.VbtName = function.FullFunctionName;
                    harvestPostCheckRow.ArgList = function.Parameters;
                    harvestPostCheckRow.Args = function.ArgList;
                    instances.Add(harvestPostCheckRow);
                }
            }
            return instances;
        }

        private void SetReadFuseArgumentCs(ref Function function)
        {
            function.SetParamValue("bankNames",
        string.Join(";", _truthTable.Rows.First().ReadFuse.Select(x => x.Field.ToUpper().Replace("CFG", "Config"))));
            function.SetParamValue("fieldNames",
                string.Join(";", _truthTable.Rows.First().ReadFuse.Select(x => x.GetAddress())));
            function.SetParamValue("flagNames", string.Join(";", _truthTable.GetHarvestEfuseRead()));
        }

        private void SetReadFuseArgumentVbt(ref Function function)
        {
            function.SetParamValue("FuseBlockName",
        string.Join(";", _truthTable.Rows.First().ReadFuse.Select(x => x.Field)));
            function.SetParamValue("FuseCategoryName",
                string.Join(";", _truthTable.Rows.First().ReadFuse.Select(x => x.GetAddress())));
            function.SetParamValue("FailFlagOrValue", string.Join(";", _truthTable.GetHarvestEfuseRead()));
        }

        private void SetWriteFuseArgumentCs(ref Function function, List<Fusing> fusing)
        {
            function.SetParamValue("bankNames", string.Join(";", fusing.Select(x => x.Field.ToUpper().Replace("CFG", "Config"))));
            function.SetParamValue("fieldNames", string.Join(";", fusing.Select(x => x.GetAddress())));
            function.SetParamValue("flagNames", string.Join(";", fusing.Select(x => x.Value)));
        }

        private void SetWriteFuseArgumentVbt(ref Function function, List<Fusing> fusing)
        {
            function.SetParamValue("FuseBlockName", string.Join(";", fusing.Select(x => x.Field)));
            function.SetParamValue("FuseCategoryName", string.Join(";", fusing.Select(x => x.GetAddress())));
            function.SetParamValue("FailFlagOrValue", string.Join(";", fusing.Select(x => x.Value)));
        }

        private SubFlowSheet CreateBinFuseFlowSheet()
        {
            string sheetName = "Flow_SetBinFuse";
            var subflow = new SubFlowSheet(sheetName);
            subflow.AddPrintStartRow(sheetName);
            int bit = _truthTable.SetBinFuseBit;
            for (int i = 0; i < bit; i++)
            {
                subflow.AddRow(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = "F_BinFuse" + i });
            }
            bool isFirst = true;
            foreach (string key in _truthTable.Rows.Find(x => x.OutputBinTables != null).OutputBinTables.Keys)
            {
                int.TryParse(key.Replace("Bin", ""), out int value);
                subflow.AddRow(new FlowRow { Opcode = isFirst ? OpCode.If : OpCode.ElseIf, Parameter = key });
                int idx = bit - 1;
                foreach (char bitValue in Convert.ToString(value, 2).PadLeft(bit, '0'))
                {
                    subflow.AddRow(new FlowRow { Opcode = bitValue.Equals('1') ? OpCode.FlagTrue : OpCode.FlagFalse, Parameter = "F_BinFuse" + idx });
                    idx--;
                }
                isFirst = false;
            }
            subflow.AddRow(new FlowRow { Opcode = OpCode.EndIf });
            subflow.AddPrintEndRow(sheetName);
            subflow.AddReturnRow();
            return subflow;
        }

        private void GenHarvestSummaryInstance(List<InstanceRow> instances, string name)
        {
            HarvestingTruthTableRow row1 = _truthTable.Rows.First();
            List<string> allFlags = _truthTable.GetHarvAllFlagList();
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
            ReOrderFlagByDomain(ref allFlags);
            var failCountRow = new InstanceRow { TestName = name };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(CSharpHarvestSummary, "scan");
            if (function.Type == ".NET")
            {
                GenerateCSharpInstanceRow(failCountRow, function, allFlags, row1);
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(HarvestSummary, "scan");
                GenerateVbtInstanceRow(failCountRow, function, allFlags, row1);
            }
            if (!instances.Exists(x => x.TestName.Equals(failCountRow.TestName, StringComparison.CurrentCultureIgnoreCase)))
            {
                instances.Add(failCountRow);
            }
        }

        private void GenerateCSharpInstanceRow(InstanceRow failCountRow, Function function, List<string> allFlags, HarvestingTruthTableRow row1)
        {
            failCountRow.VbtType = ".NET";
            failCountRow.VbtName = function.FullFunctionName;
            function.SetParamValue("allFlags", string.Join("+", allFlags.Distinct().Where(x => x != "")));
            function.SetParamValue("harvestingFlags", row1.GetHarvGlobalFailFlag());
            function.SetParamValue("sumFailCoreFlags", _truthTable.SumFailCoreFlags());
            function.SetParamValue("binoutFlags", row1.HarvBinOutFailFlag());
            failCountRow.ArgList = function.Parameters;
            failCountRow.Args = function.ArgList;
        }

        private void GenerateVbtInstanceRow(InstanceRow failCountRow, Function function, List<string> allFlags, HarvestingTruthTableRow row1)
        {
            failCountRow.VbtType = "VBT";
            failCountRow.VbtName = HarvestSummary;
            function.SetParamValue("Harv_AllFlag", string.Join("+", allFlags.Distinct().Where(x => x != "")));
            function.SetParamValue("Harv_GlobalFailFlag", row1.GetHarvGlobalFailFlag());
            function.SetParamValue("Harv_FailCoreSumFlag", _truthTable.HarvFailCoreSumFlag());
            function.SetParamValue("Harv_AllCorePassFlag", row1.HarvAllCorePassFlag());
            function.SetParamValue("Harv_BinOutFailFlag", row1.HarvBinOutFailFlag());
            function.SetParamValue("CustHarv_GlobalFailFlag", row1.GetCustHarvGlobalFailFlag());
            function.SetParamValue("CustHarv_FailCoreSumFlag", _truthTable.CustHarvFailCoreSumFlag());
            function.SetParamValue("CustHarv_BinOutFailFlag", row1.CustHarvBinOutFailFlag());
            failCountRow.ArgList = function.Parameters;
            failCountRow.Args = function.ArgList;
        }

        private void GenRunBinOutCalcScanMbistInstance(List<InstanceRow> instances, string name)
        {
            if (EpWorkbook.TestPlanWorkbook.Worksheets.Any(x => x.Name.Equals("BinOutCalcScanMbistTable", StringComparison.OrdinalIgnoreCase)))
            {
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName("FailFlagsCalcMain", "scan");
                if (function.IsFound)
                {
                    var instance = new InstanceRow { TestName = name, VbtType = function.Type, VbtName = function.FullFunctionName };

                    function.SetParamValue("sheetName", "BinOutCalcScanMbistTable");
                    instance.ArgList = function.Parameters;
                    instance.Args = function.ArgList;
                    instances.Add(instance);
                }
            }
        }

        private void ReOrderFlagByDomain(ref List<string> flags)
        {
            var flagTmp = new List<string>();
            List<string> gfxFlags = flags.FindAll(x => x.ContainsIgnoreCase("GFX"));
            List<string> socFlags = flags.FindAll(x => x.ContainsIgnoreCase("SOC"));
            List<string> pcpuFlags = flags.FindAll(x => x.ContainsIgnoreCase("PCPU"));
            List<string> ecpuFlags = flags.FindAll(x => x.ContainsIgnoreCase("ECPU"));
            List<string> binFlags = flags.FindAll(x => x.ContainsIgnoreCase("BIN"));
            flagTmp.AddRange(gfxFlags);
            flagTmp.AddRange(socFlags);
            flagTmp.AddRange(pcpuFlags);
            flagTmp.AddRange(ecpuFlags);
            foreach (string flag in flags.Where(flag => !binFlags.Exists(x => x.Equals(flag))).Where(flag => !flagTmp.Exists(x => x.Equals(flag))))
            {
                flagTmp.Add(flag);
            }
            flagTmp.AddRange(binFlags);
            flags = flagTmp;
        }

        private List<SubFlowSheet> GenHarvestFlows()
        {
            var subFlowSheets = new List<SubFlowSheet>();
            var allJobStage = TestPlanStatic.HarvestingTruthTableSheets.SelectMany(x => x.Stage).ToList();

            foreach (string jobStage in _truthTable.Stage)
            {
                string sheetName = Combination.CombineByUnderLine(new List<string> { FlowHarvestingDecisionBlock, jobStage });
                var subFlowSheet = new SubFlowSheet(sheetName, "HarvestingTruthTable");
                List<string> ignoreList = _truthTable.Job.Equals("CP1", StringComparison.OrdinalIgnoreCase) ? _truthTable.Rows.First().Flags.Where(x => Regex.IsMatch(x.Name, @"^(F_PassBinCut_|F_PWRBIN_)\w+", RegexOptions.IgnoreCase)).Select(x => x.Name).ToList() : new List<string>();
                foreach (string key in _truthTable.Rows.Find(x => x.OutputBinTables != null).OutputBinTables.Keys)
                {
                    var resultRow = new FlowRow { Opcode = OpCode.FlagFalse, Parameter = key };
                    subFlowSheet.AddRow(resultRow);
                }
                foreach (string condition in _truthTable.Rows.Select(x => x.Condition).Where(x => x != ""))
                {
                    var resultRow = new FlowRow { Opcode = OpCode.FlagFalse, Parameter = "F_" + condition };
                    subFlowSheet.AddRow(resultRow);
                }

                if (TestPlanStatic.FlagOperationSheets.Any())
                {
                    foreach (FlagOperationSheet sheet in TestPlanStatic.FlagOperationSheets)
                    {
                        if (!sheet.SubFlowName.EndsWith("_post", StringComparison.CurrentCultureIgnoreCase))
                        {
                            subFlowSheet.AddRow(new FlowRow { Opcode = OpCode.Call, Parameter = sheet.SubFlowName });
                        }
                    }
                }

                if (EpWorkbook.TestPlanWorkbook.Worksheets.Any(x => x.Name.Equals("BinOutCalcScanMbistTable", StringComparison.OrdinalIgnoreCase)))
                {
                    var runBinOutCalcScanMbistFlow = new FlowRow { Opcode = OpCode.Test, Parameter = RunBinOutCalcScanMbist };
                    subFlowSheet.AddRow(runBinOutCalcScanMbistFlow);
                }

                var harvestSummaryBeforeRow = new FlowRow { Opcode = OpCode.Test, Parameter = HarvestSummaryBeforeHarvestingDecision + "_" + jobStage.Split('_').Last() };
                subFlowSheet.AddRow(harvestSummaryBeforeRow);

                if (TestPlanStatic.FlagOperationSheets.Any())
                {
                    foreach (FlagOperationSheet sheet in TestPlanStatic.FlagOperationSheets)
                    {
                        if (sheet.SubFlowName.EndsWith("_post", StringComparison.CurrentCultureIgnoreCase))
                        {
                            subFlowSheet.AddRow(new FlowRow { Opcode = OpCode.Call, Parameter = sheet.SubFlowName });
                        }
                    }
                }

                var sumIfRow = new FlowRow { Opcode = "If", Parameter = string.Join("&&", _truthTable.Rows.SelectMany(i => i.Flags.FindAll(x => x.FlagName.Contains("_SUM"))).Select(x => "!" + x.FlagName).Distinct()) };
                subFlowSheet.AddRow(sumIfRow);

                for (int i = 0; i < _truthTable.Rows.Count; i++)
                {
                    HarvestingTruthTableRow row = _truthTable.Rows[i];
                    var ifFlowRow = new FlowRow { Opcode = "ElseIf" };
                    if (i == 0)
                    {
                        ifFlowRow.Opcode = "If";
                    }

                    List<string> ifs = BuildRowIfConditions(row, jobStage, allJobStage, ignoreList);
                    ifFlowRow.Parameter = string.Join(" && ", ifs);
                    subFlowSheet.AddRow(ifFlowRow);

                    AddDecisionRowFlows(subFlowSheet, row, jobStage);
                }

                var elseRow = new FlowRow { Opcode = OpCode.Else };
                subFlowSheet.AddRow(elseRow);

                var other = new FlowRow { Opcode = OpCode.FlagTrue, Parameter = FHarvestOtherCombination };
                subFlowSheet.AddRow(other);

                var binTable = new FlowRow { Opcode = OpCode.BinTable, Parameter = BinHarvestOtherCombination };
                subFlowSheet.AddRow(binTable);

                var endIf = new FlowRow { Opcode = "EndIf" };
                subFlowSheet.AddRow(endIf);

                subFlowSheet.AddRow(elseRow);
                subFlowSheet.AddRow(other);
                subFlowSheet.AddRow(binTable);
                subFlowSheet.AddRow(endIf);


                if (_truthTable.ReadBinFuseBit != -1 && jobStage.Contains("RD"))
                {
                    subFlowSheet.AddRows(GenerateBinFuseCheckRows());
                }

                var harvestSummaryAfterRow = new FlowRow { Opcode = OpCode.Test, Parameter = HarvestSummaryAfterHarvestingDecision + "_" + jobStage.Split('_').Last() };
                subFlowSheet.AddRow(harvestSummaryAfterRow);
                subFlowSheet.AddReturnRow();
                subFlowSheet.JobNames = new List<string> { _truthTable.Job };
                subFlowSheets.Add(subFlowSheet);
            }
            return subFlowSheets;
        }

        private List<string> BuildRowIfConditions(HarvestingTruthTableRow row, string jobStage, List<string> allJobStage, List<string> ignoreList)
        {
            var ifs = new List<string>();
            foreach (Flag flag in row.Flags)
            {
                if (flag.Value.Equals("X", StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(flag.EnableStage))
                {
                    int currentIdx = allJobStage.FindIndex(x => x.Equals(jobStage, StringComparison.OrdinalIgnoreCase));
                    int enableIdx = allJobStage.FindIndex(x => x.Equals(flag.EnableStage, StringComparison.OrdinalIgnoreCase));
                    if (currentIdx < enableIdx)
                    {
                        continue;
                    }
                }
                else if (ignoreList.Exists(x => x.Equals(flag.FlagName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                if (flag.IsSum || flag.IsCust)
                {
                    string name = flag.FlagName + "_" + flag.Value;
                    ifs.Add(name);
                }
                else
                {
                    string name = flag.FlagName;
                    if (flag.Value == "0")
                    {
                        name = "!" + name;
                    }

                    ifs.Add(name);
                }
            }
            return ifs;
        }

        private void AddDecisionRowFlows(SubFlowSheet subFlowSheet, HarvestingTruthTableRow row, string jobStage)
        {
            #region HarvestDecisionBin
            var decisionFlagRow = new FlowRow { Opcode = OpCode.FlagTrue, Parameter = row.GetDecisionResult() };
            subFlowSheet.AddRow(decisionFlagRow);

            subFlowSheet.AddRow(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = "F_" + row.Condition });
            #endregion

            #region Output Flags

            foreach (KeyValuePair<string, string> flag in row.OutputFlags)
            {
                string condition = flag.Value.Equals("1") ? OpCode.FlagTrue : OpCode.FlagFalse;
                var flagRow = new FlowRow { Opcode = condition, Parameter = flag.Key };
                subFlowSheet.AddRow(flagRow);
            }

            #endregion
            #region OutputHarvestDecisionFlags

            foreach (KeyValuePair<string, string> flag in row.OutputHarvestDecisionFlags)
            {
                if (flag.Value.Equals("1") || flag.Value.Equals("b1"))
                {
                    var flagRow = new FlowRow { Opcode = OpCode.FlagTrue, Parameter = flag.Key };
                    subFlowSheet.AddRow(flagRow);
                }
            }

            #endregion

            #region Fusing
            var addedFlags = new List<string>();
            foreach (Fusing flag in row.Fusings)
            {
                if (flag.Value.Equals("1") || flag.Value.Equals("b1", StringComparison.CurrentCulture))
                {
                    string addressFlag = flag.GetAddressFlag();
                    if (!string.IsNullOrEmpty(addressFlag) && !addedFlags.Exists(x => x.Equals(addressFlag, StringComparison.OrdinalIgnoreCase)))
                    {
                        var flagRow = new FlowRow { Opcode = OpCode.FlagTrue, Parameter = addressFlag };
                        addedFlags.Add(addressFlag);
                        subFlowSheet.AddRow(flagRow);
                    }
                }
            }
            if (row.Fusings.Any())
            {
            }

            #endregion

            #region eFuse_Write
            if (jobStage.Contains("BC"))
            {
                if (_truthTable.SetBinFuseBit != -1)
                {
                    subFlowSheet.AddRow(new FlowRow { Opcode = OpCode.Call, Parameter = "Flow_SetBinFuse" });
                }

                if (!row.Fusings.Any())
                {
                    return;
                }

                var flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = TestPlanStatic.HarvestingFuseWriteSheet == null ? row.GetConditionInstacne(_truthTable.Job) : "Harvest_eFuse_Write_" + _truthTable.Job };
                subFlowSheet.AddRow(flowRow);
            }

            #endregion
        }

        private List<FlowRow> GenerateBinFuseCheckRows()
        {
            List<FlowRow> flows = new List<FlowRow>();
            int bit = _truthTable.ReadBinFuseBit;
            bool isFirst = true;
            foreach (string key in _truthTable.Rows.Find(x => x.OutputBinTables != null).OutputBinTables.Keys)
            {
                int.TryParse(key.Replace("Bin", ""), out int value);
                var flagStatus = new List<string> { key };
                int idx = bit - 1;
                foreach (char bitValue in Convert.ToString(value, 2).PadLeft(bit, '0'))
                {
                    flagStatus.Add((bitValue.Equals('1') ? "" : "!") + "F_BinFuse" + idx);
                    idx--;
                }
                flows.Add(new FlowRow { Opcode = isFirst ? OpCode.If : OpCode.ElseIf, Parameter = string.Join("&&", flagStatus) });
                isFirst = false;
            }
            flows.Add(new FlowRow { Opcode = OpCode.Else });
            flows.Add(new FlowRow { Opcode = OpCode.Print, Parameter = '"' + "********* After the decision is done, the bin number is not matching with the previous fusing value. *********" + '"' });
            flows.Add(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = FHarvestReadBackMisMatch });
            flows.Add(new FlowRow { Opcode = OpCode.BinTable, Parameter = BinHarvestReadBackMisMatch });
            flows.Add(new FlowRow { Opcode = OpCode.EndIf });
            return flows;
        }

        private List<List<T>> GetKCombinations<T>(List<T> list, int length) where T : IComparable
        {
            if (length == 1)
            {
                return list.Select(t => new List<T> { t }).ToList();
            }

            return GetKCombinations(list, length - 1)
                .SelectMany(t => list.Where(o => o.CompareTo(t.Last()) > 0), (t1, t2) => t1.Concat(new[] { t2 }).ToList()).ToList();
        }

        private List<BinTableRow> GenFailBinTableRows()
        {
            var failBinTableRows = new List<BinTableRow>();
            List<Flag> flags = _truthTable.Rows.First().Flags;
            foreach (Flag flag in flags)
            {
                if (flag.IsSum || flag.IsCust)
                {
                    var binTableRow = new BinTableRow
                    {
                        Name = flag.GetBinTableName(),
                        Op = "And",
                        ItemList = flag.FlagName,
                        Items = new List<string> { "T" }
                    };
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Harvest", "CoreOverFailing", "", binTableRow);
                    binTableRow.Sort = binNumInfo.SoftBin.ToString("G15");
                    binTableRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                    binTableRow.Result = binNumInfo.BinNumInfo.Status;
                    failBinTableRows.Add(binTableRow);
                }
            }
            return failBinTableRows;
        }

        private void GenOverFailingFlowSheet(string block)
        {
            List<Flag> flags = _truthTable.Rows.First().Flags;
            var flowRows = new List<FlowRow>();
            foreach (Flag flag in flags)
            {
                if (flag.IsSum || flag.IsCust)
                {
                    var flowRow = new FlowRow { Opcode = OpCode.BinTable, Parameter = flag.GetBinTableName() };
                    flowRows.Add(flowRow);
                }
            }

            if (flowRows.Any())
            {
                var flowSheet = new SubFlowSheet(string.IsNullOrEmpty(block) ? "Flow_CoreOverFailing" : $"Flow_CoreOverFailing_{block}");
                var summaryInstance = new FlowRow { Opcode = OpCode.Test, Parameter = string.IsNullOrEmpty(block) ? HarvestSummaryCoreOverFailing : $"{HarvestSummaryCoreOverFailing}_{block}" };

                flowSheet.AddRow(summaryInstance);
                flowSheet.AddRows(flowRows);
                flowSheet.AddReturnRow();
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirNonBinCut, flowSheet);
            }
        }

        private List<BinTableRow> GenPassBinTableRows()
        {
            var repairList = new List<string>
            {
                "F_SocMbist_Repair_check_Final",
                "F_CpuMbist_Repair_check_Final",
                "F_GfxMbist_Repair_check_Final"
            };

            if (_truthTable.GetRepairFlagList().Any())
            {
                repairList = _truthTable.GetRepairFlagList();
            }

            var binTableRows = new List<BinTableRow>();
            var mainBinTableRows = new List<BinTableRow>();
            var passBinTableRows = new List<BinTableRow>();

            foreach (HarvestingTruthTableRow row in _truthTable.Rows)
            {
                if (!row.ProposedBinName.ContainsIgnoreCase("REPAIR["))
                {
                    repairList = new List<string>();
                }

                GenBinTableRowEx(row, passBinTableRows, repairList);
            }
            mainBinTableRows = mainBinTableRows.OrderBy(x => x.Name).ToList();

            binTableRows.AddRange(mainBinTableRows);
            binTableRows.AddRange(passBinTableRows);
            return binTableRows;
        }

        private void GenBinTableRowEx(HarvestingTruthTableRow row, List<BinTableRow> passBinTableRows, List<string> repairList)
        {
            string name = "";
            foreach (KeyValuePair<string, string> outputFlag in row.OutputBinTables)
            {
                if (outputFlag.Value.Equals("1"))
                {
                    name = outputFlag.Key;//Combination.CombineByUnderLine(outputFlag.Key, _truthTable.Job);
                    break;
                }
            }

            var dic = new Dictionary<string, string>();
            var combinationList = new List<List<List<string>>>();
            var allFlags = new List<string>();
            foreach (Flag flag in row.Flags)
            {
                if (!Regex.IsMatch(flag.Value, @"\d+"))
                {
                    continue;
                }

                string flagName = flag.FlagName;
                if (flag.IsSum || flag.IsCust)
                {
                    dic.Add(flagName + "_" + flag.Value, "T");
                    int.TryParse(flag.Value, out int value);
                    allFlags.Add(flagName + "_" + flag.Value);
                }
                else
                {
                    if (flag.Value.Equals("1"))
                    {
                        dic.Add(flagName, "T");
                    }
                    else if (flag.Value.Equals("0"))
                    {
                        dic.Add(flagName, "F");
                    }

                    allFlags.Add(flagName);
                }
            }


            if (!string.IsNullOrEmpty(name))
            {
                //Combinations
                List<List<List<string>>> finalResult = Combination.GetAllPossibleCombos(combinationList);
                for (int index = 0; index < finalResult.Count; index++)
                {
                    var combination = finalResult[index].SelectMany(x => x).ToList();
                    passBinTableRows.AddRange(GenPassBinTableRow(name, row, allFlags, repairList,
                        combination, dic, ref ExistHarvBinTable));
                }
            }
        }

        private List<BinTableRow> GenPassBinTableRow(string name, HarvestingTruthTableRow row, List<string> flags, List<string> repairList, List<string> combination, Dictionary<string, string> dic, ref Dictionary<string, string> existHarvBinTable)
        {
            var passBinTableRows = new List<BinTableRow>();
            Dictionary<string, Dictionary<string, string>> proposedBinNameList = row.GetProposedBinNameList();
            double repairCount = Math.Pow(2, repairList.Count);
            if (proposedBinNameList == null)
            {
                for (int j = 0; j < repairCount; j++) //repairList
                {
                    var binTableHavestRow = new BinTableRow();
                    string rowName = Combination.CombineByUnderLine(new List<string>
                    {
                        name,
                        row.GetModifiedCondition(),
                    }).Replace(" ", "_");

                    binTableHavestRow.Name = rowName.Trim('_');

                    var failFlags = new List<string>();
                    failFlags.AddRange(flags);
                    failFlags.AddRange(repairList);
                    binTableHavestRow.ItemList = string.Join(",", failFlags);
                    GetBinItemsByTruthTable(j, binTableHavestRow, flags, repairList, combination, dic);
                    binTableHavestRow.Op = "AND";
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Harvest", row.Condition, "", binTableHavestRow);
                    binTableHavestRow.Result = binNumInfo.BinNumInfo.Status;
                    binTableHavestRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                    binTableHavestRow.Sort = binNumInfo.SoftBin.ToString("G15");
                    passBinTableRows.Add(binTableHavestRow);
                }
            }
            else
            {
                foreach (KeyValuePair<string, Dictionary<string, string>> proposedBinName in proposedBinNameList)
                {
                    string existSortBinNum = "";
                    for (int j = 0; j < repairCount; j++) //repairList
                    {
                        var binTableHavestRow = new BinTableRow();
                        string formatNum = Convert.ToString(j, 2);
                        formatNum = byte.Parse(formatNum).ToString("D" + repairList.Count);
                        string nameTmp = proposedBinName.Key;
                        string rowName = nameTmp.Replace("repair", "repair" + formatNum);

                        binTableHavestRow.Name = rowName.Trim('_');
                        if (existHarvBinTable.TryGetValue(binTableHavestRow.Name, out string value))
                        {
                            existSortBinNum = value;
                        }
                        var flagTmp = new List<string>();
                        flagTmp.AddRange(flags);
                        var dicTmp = new Dictionary<string, string>();
                        foreach (KeyValuePair<string, string> item in dic)
                        {
                            dicTmp.Add(item.Key, item.Value);
                        }

                        if (proposedBinName.Value.Count != 0)
                        {
                            foreach (KeyValuePair<string, string> item in proposedBinName.Value)
                            {
                                dicTmp.Add(item.Key, item.Value);
                            }
                            flagTmp.AddRange(proposedBinName.Value.Keys);
                        }
                        var failFlags = new List<string>();
                        failFlags.AddRange(flagTmp);
                        failFlags.AddRange(repairList);
                        binTableHavestRow.ItemList = string.Join(",", failFlags);
                        GetBinItemsByTruthTable(j, binTableHavestRow, flagTmp, repairList, combination, dicTmp);
                        binTableHavestRow.Op = "AND";
                        BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Harvest", row.Condition, "", binTableHavestRow);
                        binTableHavestRow.Result = binNumInfo.BinNumInfo.Status;
                        binTableHavestRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                        binTableHavestRow.Sort = binNumInfo.SoftBin.ToString("G15");
                        if (existSortBinNum == "")
                        {
                            binTableHavestRow.ColumnA = row.Condition;
                            binTableHavestRow.Sort = binNumInfo.SoftBin.ToString("G15");
                            existHarvBinTable.Add(binTableHavestRow.Name, binTableHavestRow.Sort);
                        }
                        else
                        {
                            binTableHavestRow.Sort = existSortBinNum;
                            binTableHavestRow.ColumnA = row.Condition + ";Duplicate BinTable";
                            binTableHavestRow.Name = binTableHavestRow.Name + "_" + row.Condition;
                        }
                        passBinTableRows.Add(binTableHavestRow);
                    }
                }
            }
            return passBinTableRows;
        }



        private void GetBinItemsByTruthTable(int j, BinTableRow row, List<string> flags, List<string> repairList, List<string> combinations, Dictionary<string, string> dic)
        {
            #region Harvest
            foreach (string flag in flags)
            {
                if (dic.TryGetValue(flag, out string value))
                {
                    row.Items.AddRange(new List<string> { value });
                }
                else
                {
                    if (combinations.Exists(y => y.Equals(flag, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        row.Items.AddRange(new List<string> { "T" });
                    }
                    else
                    {
                        row.Items.AddRange(new List<string> { "F" });
                    }
                }
            }
            #endregion

            #region repair
            if (repairList.Any())
            {
                string formatNum = Convert.ToString(j, 2);
                formatNum = byte.Parse(formatNum).ToString("D" + repairList.Count);
                foreach (char c in formatNum)
                {
                    if (c.Equals('1'))
                    {
                        row.Items.Add("T");
                    }
                    else
                    {
                        row.Items.Add("F");
                    }
                }

            }
            #endregion
        }

        private void TransferValue(string bitValue, List<string> flagList, ref Dictionary<string, string> flagDic)
        {
            int index = 0;
            if (bitValue.StartsWith("b") && int.TryParse(bitValue.TrimStart('b'), out _))
            {
                foreach (char bit in bitValue.Replace("b", ""))
                {
                    flagDic.Add(flagList[index], bit == '0' ? "F" : "T");
                    index++;
                }
            }
            else if (bitValue.StartsWith("x"))
            {
                if (int.TryParse(bitValue.Replace("x", ""), out int value))
                {
                    foreach (char bit in Convert.ToString(Convert.ToInt32(value.ToString(), 16), 2).PadLeft(flagList.Count, '0'))
                    {
                        flagDic.Add(flagList[index], bit == '0' ? "F" : "T");
                        index++;
                    }
                }
            }
        }

        public SubFlowSheet GenHarvPostCheckSubFlow(List<InstanceRow> instanceRows)
        {
            string sheetName = "Flow_HarvestPostCheck";
            var subFlowSheet = new SubFlowSheet(sheetName) { JobNames = new List<string>() };
            foreach (InstanceRow harvestInst in instanceRows.Where(x => x.VbtName.Split('.').LastOrDefault() == CheckFromFuseToFlag))
            {
                var flowRow = new FlowRow();
                string job = harvestInst.TestName.Split('_').Last();
                flowRow.Opcode = OpCode.Test;
                flowRow.Job = job;
                flowRow.Parameter = harvestInst.TestName;
                subFlowSheet.AddRow(flowRow);
                if (!subFlowSheet.JobNames.Exists(x => x.Equals(job)))
                {
                    subFlowSheet.JobNames.Add(job);
                }
            }
            var row = new FlowRow { Opcode = OpCode.BinTable, Parameter = BinHarvestPostCheckFail };
            subFlowSheet.AddRow(row);
            subFlowSheet.AddReturnRow();

            return subFlowSheet;
        }

        public SubFlowSheet GenHarvReadSubFlow(List<InstanceRow> instanceRows)
        {
            string sheetName = "Flow_Harvest_eFuse_Read";
            var subFlowSheet = new SubFlowSheet(sheetName) { JobNames = new List<string>() };
            if (TestPlanStatic.HarvestingFlagInitSheet != null)
            {
                Dictionary<string, string> bankDic = TestPlanStatic.HarvestingFlagInitSheet.BankDictionary;
                if (TestPlanStatic.HarvestingFlagInitSheet.EfuseMappingTable != null)
                {
                    foreach (HarvestingFuseRead row in TestPlanStatic.HarvestingFlagInitSheet.EfuseMappingTable)
                    {
                        var flagInitDic = new Dictionary<string, string>();
                        var flagTrueList = new List<string>();
                        var flagFalseList = new List<string>();
                        foreach (KeyValuePair<string, string> fuseRow in row.FuseMapping)
                        {
                            string fieldName = fuseRow.Value;
                            string flagName = fuseRow.Key;
                            if (!bankDic.ContainsKey(fieldName))
                            {
                                if (_regex2.IsMatch(flagName))
                                {
                                    string flag = _regex2.Match(flagName).Groups["flag"].ToString();
                                    _regex4.Replace(flag, "");
                                    GroupItem iterates = flag.GetIterates();
                                    var flagList = new List<string>();

                                    foreach (KeyValuePair<int, List<Tuple<string, string>>> iterate in iterates.Groups)
                                    {
                                        string text = flag;
                                        foreach (Tuple<string, string> item in iterate.Value)
                                        {
                                            text = text.Replace(item.Item1, item.Item2);
                                        }

                                        flagList.Add(text);
                                    }
                                    if (!iterates.Groups.Any())
                                    {
                                        flagList.Add(flagName);
                                    }

                                    TransferValue(fieldName, flagList, ref flagInitDic);
                                }
                            }
                        }
                        foreach (KeyValuePair<string, string> flagInit in flagInitDic)
                        {
                            string status = flagInit.Value;
                            if (status.Equals("T"))
                            {
                                flagTrueList.Add(flagInit.Key);
                            }
                            else
                            {
                                flagFalseList.Add(flagInit.Key);
                            }
                        }
                        if (flagFalseList.Any())
                        {
                            subFlowSheet.AddRow(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = string.Join(",", flagFalseList), Job = row.Job });
                        }

                        if (flagTrueList.Any())
                        {
                            subFlowSheet.AddRow(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = string.Join(",", flagTrueList), Job = row.Job });
                        }
                    }
                }
            }
            foreach (InstanceRow harvestInst in instanceRows)
            {
                string functionName = harvestInst.VbtName.Split('.').LastOrDefault();
                if (functionName != HarvestEfuseRead && functionName != CopyFromFuseToFlag)
                {
                    continue;
                }

                var flowRow = new FlowRow();
                string job = harvestInst.TestName.Split('_').Last();
                flowRow.Opcode = OpCode.Test;
                flowRow.Job = job;
                flowRow.Parameter = harvestInst.TestName;
                subFlowSheet.AddRow(flowRow);
                if (!subFlowSheet.JobNames.Exists(x => x.Equals(job)))
                {
                    subFlowSheet.JobNames.Add(job);
                }
            }
            //Add HarvestSummaryAfterInit instance
            var summaryAfterInitRow = new FlowRow { Parameter = HarvestSummaryAfterInit, Opcode = OpCode.Test };
            subFlowSheet.AddRow(summaryAfterInitRow);

            subFlowSheet.AddReturnRow();

            return subFlowSheet;
        }
    }
}
