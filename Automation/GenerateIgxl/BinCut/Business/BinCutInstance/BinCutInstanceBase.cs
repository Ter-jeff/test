using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Atpg;

using CommonLib.Enums;
using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.BinCut.Business.BinCutInstance
{
    public abstract class BinCutInstanceBase : IBinCutInstance
    {
        public BinCutFinalInstanceRow BinCutFinalInstanceRow;
        protected BinCutSourceItem SourceRow;
        protected readonly BinCutInputData BinCutInputManager;
        protected string PerformanceMode;             //Performance mode from BinCut Flow
        protected string Module;                      //CPU OR GFX OR SOC OR DDR 
        protected const string Min = "Min";
        protected const string Max = "Max";
        protected const string Typ = "Typ";
        protected string Block;

        protected BinCutInstanceBase(BinCutFinalInstanceRow binCutFinalInstanceRow, BinCutSourceItem sourceRow, BinCutInputData binCutInputManager)
        {
            BinCutFinalInstanceRow = binCutFinalInstanceRow;
            SourceRow = sourceRow;
            BinCutInputManager = binCutInputManager;
        }

        public virtual HashSet<string> AllJobs
        {
            get
            {
                if (BinCutInputManager.NewBinCutFlowTables != null && BinCutInputManager.NewBinCutFlowTables.Any())
                {
                    return BinCutInputManager.NewBinCutFlowTables.Select(x => x.JobName).ToHashSet(StringComparer.OrdinalIgnoreCase);
                }

                return BinCutInputManager.BinCutFlowTables.Select(x => x.JobName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
        }

        public InstanceRow GenerateInstance()
        {
            var row = new InstanceRow { TestName = GetInstanceName() };
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                string vbtName = GetVbtNameCs();
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "binCut", true);
                if (!function.IsFound || function.Type == "VBT")
                {
                    row.VbtName = GetVbtName();
                    row.VbtType = "VBT";
                }
                else
                {

                    row.VbtName = function.FullFunctionName;
                    row.VbtType = ".NET";
                }
            }
            else
            {
                row.VbtName = GetVbtName();
                row.VbtType = "VBT";
            }
            row.InitList = BinCutFinalInstanceRow.InitList;
            row.PayloadList = BinCutFinalInstanceRow.PayloadList;
            string selector = BinCutFinalInstanceRow.GetVoltageType();
            if (!BinCutFinalInstanceRow.Nop && !(!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) && SourceRow.ColumnName.Equals(EnumColumnName.FUNC)) && !(SourceRow.InstOrCallFlowByBms && SourceRow.ColumnName != EnumColumnName.E1Voltage))
            {
                GenerateArgsAndArgList(row);
                if (!row.VbtName.Equals("Set_E1_Voltage_ForPmode", StringComparison.CurrentCultureIgnoreCase) &&
                    !row.VbtName.Split('.').Last().Equals("SetVoltageWithoutTest", StringComparison.CurrentCultureIgnoreCase))
                {
                    row.TimeSets = BinCutFinalInstanceRow.GetTimeSetVersion(BinCutFinalInstanceRow.PayloadList);
                    row.AcCategory = string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.AcSpec) ? GenerateAcCategory(row) : BinCutFinalInstanceRow.BinCutInstanceRow.AcSpec;
                    row.AcSelector = GenerateAcSelector();
                    row.DcCategory = GenerateDcCategory();
                    row.DcSelector = GenerateDcSelector(selector);
                    row.PinLevels = GenerateLevel();
                }
            }
            if (BinCutFinalInstanceRow.Nop)
            {
                row.ColumnA = "BlankInstance";
            }
            else if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) && SourceRow.ColumnName.Equals(EnumColumnName.FUNC))
            {
                row.ColumnA = "UnknownTestType";
            }
            else
            {
                row.ColumnA = GetColumnA(SourceRow.ColumnName, BinCutFinalInstanceRow.BinCutInstanceRow);
            }

            row.ColumnA += ";DC Category:" + row.DcCategory;
            row.RowNum = BinCutFinalInstanceRow.BinCutInstanceRow.RowNum;
            row.FinalJobs = BinCutFinalInstanceRow.FinalJobs;
            return row;
        }

        protected virtual string GetFlag()
        {
            string flagName = "";
            switch (SourceRow.ColumnName)
            {
                case EnumColumnName.TD:
                    flagName = "F_TD_" + SourceRow.TargetPerformanceMode + "_HBV";
                    break;
                case EnumColumnName.Mbist:
                    flagName = "F_Mbist_" + SourceRow.TargetPerformanceMode + "_HBV";
                    break;
                case EnumColumnName.FUNC:
                    flagName = "F_RTOS_" + SourceRow.TargetPerformanceMode + "_HBV";
                    break;
                case EnumColumnName.ELB:
                    flagName = "F_ELB_" + SourceRow.TargetPerformanceMode + "_HBV";
                    break;
                case EnumColumnName.ILB:
                    flagName = "F_ILB_" + SourceRow.TargetPerformanceMode + "_HBV";
                    break;
            }
            return flagName;
        }

        protected virtual string GetFlagBv()
        {
            string flagName = "";
            switch (SourceRow.ColumnName)
            {
                case EnumColumnName.TD:
                    flagName = "F_" + SourceRow.TargetPerformanceMode + "_TD_BV";
                    break;
                case EnumColumnName.Mbist:
                    flagName = "F_" + SourceRow.TargetPerformanceMode + "_Mbist_BV";
                    break;
                case EnumColumnName.FUNC:
                    flagName = "F_" + SourceRow.TargetPerformanceMode + "_RTOS_BV";
                    break;
                case EnumColumnName.ELB:
                    flagName = "F_" + SourceRow.TargetPerformanceMode + "_ELB_BV";
                    break;
                case EnumColumnName.ILB:
                    flagName = "F_" + SourceRow.TargetPerformanceMode + "_ILB_BV";
                    break;
            }
            return flagName;
        }

        internal void SetVbtArg(ref InstanceRow row, BinCutFinalInstanceRow binCutFinalInstanceRow)
        {
            string vbtName = row.VbtName;
            string dsscPat = "";
            if (vbtName.ContainsIgnoreCase("callinstance"))
            {
                string name = row.TestName;
                if (row.DcSelector == Min)
                {
                    name = Regex.Replace(name, "_NV$", "_LV", RegexOptions.IgnoreCase);
                }
                else if (row.DcSelector == Max)
                {
                    name = Regex.Replace(name, "_NV$", "_HV", RegexOptions.IgnoreCase);
                }

                row.SetArgument("inst_CallInstance", name);
            }
            if (binCutFinalInstanceRow.PatternList.Any())
            {
                foreach (string pat in binCutFinalInstanceRow.PatternList)
                {
                    if (Regex.IsMatch(pat, @"\w*DSSC\w*", RegexOptions.IgnoreCase))
                    {
                        dsscPat = pat;
                    }
                }
            }
            if (!string.IsNullOrEmpty(dsscPat))
            {
                string sendPinName = "JTAG_TDI";
                if (LocalSpecs.HardIpInfos != null)
                {
                    HardIpInfo target = LocalSpecs.HardIpInfos.GetHardIpInfo(dsscPat);
                    sendPinName = string.IsNullOrEmpty(target.SendPinName) ? sendPinName : target.SendPinName;
                }
                row.SetArgument("DigSrc_Pin", sendPinName);
            }
            row.SetArgument("Performance_mode", GetBinningDomain());
        }

        internal void SetCsArg(ref InstanceRow row)
        {
            string name = row.TestName;
            bool isHbv = SourceRow.TableType.ToString() == "Hv";
            bool isPost = SourceRow.TableType.ToString() == "Post";
            string failFlag = isHbv ? GetFlag() : GetFlagBv();
            if (isPost)
            {
                failFlag = GetFlagPost();
            }

            if (row.DcSelector == Min)
            {
                name = Regex.Replace(name, "_NV$", "_LV", RegexOptions.IgnoreCase);
            }
            else if (row.DcSelector == Max)
            {
                name = Regex.Replace(name, "_NV$", "_HV", RegexOptions.IgnoreCase);
            }

            row.SetArgument("performanceMode", (isHbv ? "HBV_" : "") + SourceRow.PerformanceMode);
            row.SetArgument("callInstanceName", name);
            row.SetArgument("instanceFailFlag", failFlag);
        }

        private string GetFlagPost()
        {
            return $"F_{SourceRow.ColumnName}_{SourceRow.PerformanceMode}_outsidebincut_BV";
        }

        public List<InstanceRow> GenerateInstanceByTestName()
        {
            string vbtName;
            Function function;
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                vbtName = GetVbtNameCs();
                function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "binCut", true);
                if (!function.IsFound || function.Type == "VBT")
                {
                    vbtName = GetVbtName();
                    function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "binCut");
                }
            }
            else
            {
                vbtName = GetVbtName();
                function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "binCut");
            }


            #region Hardip
            var instanceRows = new List<InstanceRow>();
            if (BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Hardip)
            {

                var dic = new Dictionary<string, string>
                {
                    { "PATSET", "DQSSWPPAT" },
                    { "DIGCAP_DATAWIDTH","NOOFBISTS" },
                    { "DIGCAP_SAMPLE_SIZE", "DIGCAP_SAMPLE_SIZE_DQS" },
                    { "CUS_STR_MAINPROGRAM","SWEEPVTSTR" } ,
                    { "CUS_STR_DIGCAPDATA","CUS_STR_DIGCAPDATA_DQS" }
                };
                var instanceRow = new InstanceRow();
                string selector = BinCutFinalInstanceRow.GetVoltageType();
                instanceRow.VbtName = function.FullFunctionName;
                instanceRow.VbtType = function.Type;
                if (selector.Equals("UnknowType"))
                {
                    selector = GenerateDcSelector(selector);
                    switch (selector)
                    {
                        case "Min":
                            selector = "LV";
                            break;
                        case "Typ":
                            selector = "NV";
                            break;
                        case "Max":
                            selector = "HV";
                            break;
                    }
                }
                if (BinCutInputManager.HardIpPatterns != null)
                {
                    List<InstanceRow> rows = BinCutInputManager.GenHardipInstanceByPattern(BinCutFinalInstanceRow.PatternList.First(), TestPlanStatic.HardIpDcSheet);
                    if (rows != null && rows.Count != 0)
                    {
                        instanceRow = rows.Find(x => x.TestName.Contains(selector));
                        if (instanceRow != null)
                        {
                            function.ReplaceFunction(instanceRow, dic);
                            instanceRow.DcCategory = GenerateDcCategory();
                            instanceRow.DcSelector = GenerateDcSelector(selector);
                            instanceRow.PayloadList = BinCutFinalInstanceRow.PatternList;
                            instanceRow.RowNum = BinCutFinalInstanceRow.BinCutInstanceRow.RowNum;
                            instanceRow.FinalJobs = BinCutFinalInstanceRow.FinalJobs;
                            instanceRow.ColumnA = BinCutFinalInstanceRow.Nop ? "BlankInstance" : GetColumnA(SourceRow.ColumnName, BinCutFinalInstanceRow.BinCutInstanceRow);
                            instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
                            instanceRow.ArgList = function.Parameters;
                            instanceRow.Args = function.ArgList;
                            if (function.Type == ".NET")
                            {
                                SetCsArg(ref instanceRow);
                            }
                            else
                            {
                                SetVbtArg(ref instanceRow, BinCutFinalInstanceRow);
                            }
                        }
                    }
                }

                if (instanceRow != null)
                {
                    instanceRow.TestName = GetInstanceName();
                    instanceRows.Add(instanceRow);
                }
            }
            return instanceRows;
            #endregion
        }

        public FlowRow GenerateFlowRow(bool isHvccOrPost, bool isTmps, bool isCsharp)
        {
            var flowRow = new FlowRow();
            if (BinCutFinalInstanceRow.BinCutInstanceRow != null && (!string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.SheetName) || SourceRow.InstOrCallFlowByBms))
            {
                flowRow.ColumnA = GetColumnA(SourceRow.ColumnName, BinCutFinalInstanceRow.BinCutInstanceRow);
            }

            if (SourceRow.ColumnName.Equals(EnumColumnName.CallNwireEnable) || SourceRow.ColumnName.Equals(EnumColumnName.CallNwireDisable) || isTmps)
            {
                SetCallFlowRow(flowRow);
            }
            else
            {
                if (SourceRow.ColumnName.Equals(EnumColumnName.RelayOn) || SourceRow.ColumnName.Equals(EnumColumnName.RelayOff))
                {
                    InstanceSheet instanceCommon = TestProgram.IgxlWorkBk.InsSheets.FirstOrDefault(x => x.Value.Name.Equals("TestInst_Common")).Value;
                    InstanceRow instanceRow = instanceCommon.Rows.FirstOrDefault(x => x.TestName.Equals(SourceRow.ColumnContent.Split(':').Last(), StringComparison.OrdinalIgnoreCase));
                    flowRow.Opcode = OpCode.Test;
                    flowRow.Parameter = SourceRow.ColumnContent.Split(':').Last();
                    flowRow.Job = BinCutFinalInstanceRow.GetJob();
                    flowRow.Enable = GetEnableWord();
                    flowRow.Env = instanceRow != null ? GetEnv() : "MissingRelayInstance";
                }
                else
                {
                    SetTestFlowRow(flowRow, isHvccOrPost, isCsharp);
                }
            }

            return flowRow;
        }

        private void SetCallFlowRow(FlowRow flowRow)
        {
            string callFlowName = SourceRow.ColumnContent.Split(':').Last();
            bool foundInCustom = false;
            SubFlowSheet targetFlowSheet = TestProgram.IgxlWorkBk.SubFlowSheets.FirstOrDefault(x => x.Value.Name.Equals(callFlowName)).Value;
            if (targetFlowSheet == null)
            {
                foreach (string customPath in LocalSpecs.CustomPath)
                {
                    if (Directory.Exists(customPath) && !foundInCustom)
                    {
                        var customSubFlowNames = IgxlSheetReaderHelpers.GetSheetsByType(customPath, EnumSheetType.DTFlowtableSheet).Select(Path.GetFileNameWithoutExtension).ToList();
                        foundInCustom = customSubFlowNames.Contains(callFlowName);
                    }
                }
            }
            flowRow.Opcode = OpCode.Call;
            flowRow.Enable = GetEnableWord();
            flowRow.Parameter = callFlowName;
            flowRow.Env = targetFlowSheet != null || foundInCustom ? "" : "MissingCallingFlowSheet";
        }

        private void SetTestFlowRow(FlowRow flowRow, bool isHvccOrPost, bool isCsharp)
        {
            flowRow.Opcode = BinCutFinalInstanceRow.NopByEnableWord ? OpCode.Nop : OpCode.Test;
            flowRow.Parameter = GetInstanceName();
            flowRow.Enable = GetEnableWord();
            GetDeviceCondition(ref flowRow);
            flowRow.Job = BinCutFinalInstanceRow.GetJob();
            if (isHvccOrPost && BinCutFinalInstanceRow.BinCutInstanceRow != null && !BinCutFinalInstanceRow.BinCutInstanceRow.FailFlag.Equals("X", StringComparison.OrdinalIgnoreCase) && !BinCutFinalInstanceRow.BinCutInstanceRow.HarvPinGrpEnable.Equals("True", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.HarvestBinningFlag) && !isCsharp)
            {
                flowRow.FailAction = GetFlag();
            }
            else if (BinCutFinalInstanceRow.BinCutInstanceRow != null && BinCutFinalInstanceRow.BinCutInstanceRow.FailFlag != null && !BinCutFinalInstanceRow.BinCutInstanceRow.FailFlag.Equals("X", StringComparison.OrdinalIgnoreCase) && isCsharp)
            {
                flowRow.FailAction = BinCutFinalInstanceRow.BinCutInstanceRow.FailFlag;
            }

            flowRow.Env = GetEnv();
        }

        private string GetEnv()
        {
            var list = new List<string>();
            if (BinCutFinalInstanceRow.Nop)
            {
                return "BlankInstance";
            }

            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) && SourceRow.ColumnName.Equals(EnumColumnName.FUNC))
            {
                return "UnknownTestType";
            }

            if (BinCutFinalInstanceRow.PatternList.Count != 0 && BinCutFinalInstanceRow.BinCutInstanceRow.Type != BincutInstanceType.Rtos && !SourceRow.InstOrCallFlowByBms)
            {
                List<string> patternList = BinCutFinalInstanceRow.PatternList;
                if (BinCutFinalInstanceRow.PatternList.Exists(x => x.Contains("+")))
                {
                    patternList = new List<string>();
                    foreach (string pat in BinCutFinalInstanceRow.PatternList)
                    {
                        if (pat.Contains("+"))
                        {
                            patternList.AddRange(pat.Split(new[] { '+', ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList());
                        }
                        else
                        {
                            patternList.Add(pat);
                        }
                    }
                }

                foreach (string pat in patternList)
                {
                    string temEnv = AcTSetCategoryMapSingleton.Instance().CheckAllPatternExist(pat);
                    if (!string.IsNullOrEmpty(temEnv))
                    {
                        list.Add(temEnv);
                    }
                }
            }

            list = list.Distinct().ToList();
            return string.Join(",", list);
        }

        private void GetDeviceCondition(ref FlowRow row)
        {
            if (BinCutFinalInstanceRow.BinCutInstanceRow.EnableAndDevice == null && BinCutFinalInstanceRow.BinCutInstanceRow == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.EnableAndDevice))
            {
                if (BinCutFinalInstanceRow.BinCutInstanceRow.EnableAndDevice != null)
                {
                    string[] arr = BinCutFinalInstanceRow.BinCutInstanceRow.EnableAndDevice.Split(',');
                    foreach (string data in arr)
                    {
                        if (string.IsNullOrEmpty(data))
                        {
                            continue;
                        }

                        string deviceName = data;
                        if (data.EndsWith("@site", StringComparison.CurrentCultureIgnoreCase))
                        {
                            deviceName = Regex.Replace(data, "@site", "", RegexOptions.IgnoreCase).Trim().TrimStart('!');
                        }

                        if (data.EndsWith("@site", StringComparison.CurrentCultureIgnoreCase))
                        {
                            row.DeviceCondition = data.StartsWith("!") ? "Flag-false" : "Flag-true";
                            row.DeviceName = deviceName;
                        }
                    }
                }
            }
        }

        private string GetEnableWord()
        {
            var jobs = BinCutFinalInstanceRow.FinalJobs.Select(x => x.Trim()).ToList();
            string job = SourceRow.GetEnableWord(BinCutFinalInstanceRow.FinalJobs);
            string word = job;
            string sourceFlowName = SourceRow.PerformanceMode.Split('_').First();
            if (SourceRow.Nop)
            {
                word = Combination.CombineEnableWord(word, "NonBinCutOrder");
            }

            string enableFlow = BinCutFinalInstanceRow.BinCutInstanceRow.EnableFlow;
            if (string.IsNullOrEmpty(enableFlow))
            {
                enableFlow = BinCutFinalInstanceRow.BinCutInstanceRow.EnableAndDevice;
            }

            if (!string.IsNullOrEmpty(enableFlow))
            {
                var enableWords = new List<string>();
                string[] arr = enableFlow.Split(',');
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

                    if (Regex.IsMatch(data, "@site$", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    if (Regex.IsMatch(data, @"^(?<enableword>.+)([\(](?<FlowName>[^)]+)[\)])$", RegexOptions.IgnoreCase))
                    {
                        Match regexCondition = Regex.Match(data, @"^(?<enableword>.+)([\(](?<FlowName>[^)]+)[\)])$");
                        string flowName = regexCondition.Groups["FlowName"].ToString();
                        string enableWord = regexCondition.Groups["enableword"].ToString();
                        if (flowName.Equals(sourceFlowName))
                        {
                            enableWords.Add(enableWord);
                        }
                    }
                    else
                    {
                        enableWords.Add(data);
                    }
                }
                if (enableWords.Except(jobs).Any())
                {
                    word = Combination.CombineEnableWord(word, string.Join("||", enableWords));
                }
            }

            return word;
        }

        public FlowRow GetBinTableRow()
        {
            var flowRow = new FlowRow { Opcode = OpCode.BinTable, Parameter = GetBinTableName() };
            return flowRow;
        }

        public FlowRow GetBinTableRowBv()
        {
            var flowRow = new FlowRow { Opcode = OpCode.BinTable, Parameter = GetBinTableNameBv() };
            return flowRow;
        }

        public string GetBinTableName()
        {
            return Regex.Replace(GetFlag(), "^F_", "Bin_");
        }

        public string GetBinTableNameBv()
        {
            return Regex.Replace(GetFlagBv(), "^F_", "Bin_");
        }

        protected string GetModuleFromInstanceName()
        {
            if (BinCutFinalInstanceRow.PatSetName == "")
            {
                return SourceRow.GetDomainOfContent();
            }

            List<string> nameList = string.IsNullOrEmpty(BinCutFinalInstanceRow.InitPatSetName) ?
                BinCutFinalInstanceRow.PatSetName.Split('_').ToList()
                : BinCutFinalInstanceRow.InitPatSetName.Split('_').ToList();
            string first = nameList.First();
            if (first.StartsWith("Cpu", StringComparison.CurrentCultureIgnoreCase))
            {
                return BinCutConstant.ConCpu;
            }

            if (first.StartsWith("Gpu", StringComparison.CurrentCultureIgnoreCase) ||
                first.StartsWith("Gfx", StringComparison.CurrentCultureIgnoreCase))
            {
                return BinCutConstant.ConGpu;
            }

            if (first.StartsWith("Soc", StringComparison.CurrentCultureIgnoreCase))
            {
                return BinCutConstant.ConSoc;
            }

            string result = Regex.Replace(first, "TD|Mbist", "", RegexOptions.IgnoreCase);
            return result;
        }

        protected void ExecuteCoreLogic(InstanceRow row,
            Action<Function, InstanceRow, string> generateCSharpInstanceRow,
            Action<Function, InstanceRow, string> generateVbtInstanceRow)
        {
            string vbtName;
            Function function;
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                vbtName = GetVbtNameCs();
                function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "binCut", true);
                row.VbtName = function.FullFunctionName;
                if (!function.IsFound || function.Type == "VBT")
                {
                    vbtName = GetVbtName();
                    function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "binCut");
                    row.VbtName = function.FunctionName;
                }
            }
            else
            {
                vbtName = GetVbtName();
                function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "binCut");
                row.VbtName = function.FunctionName;
            }

            if (function.Type == ".NET")
            {
                generateCSharpInstanceRow(function, row, vbtName);
            }
            else
            {
                generateVbtInstanceRow(function, row, vbtName);
            }
        }

        protected virtual void GenerateArgsAndArgList(InstanceRow row)
        {
            ExecuteCoreLogic(row,
                GenerateCSharpInstanceRow,
                GenerateVbtInstanceRow);
        }

        private void GenerateCSharpInstanceRow(Function function, InstanceRow row, string vbtName)
        {
            bool isHbv = SourceRow.TableType.ToString() == "Hv";
            if (vbtName.EndsWith("BinCutTest", StringComparison.CurrentCultureIgnoreCase))
            {
                function.SetParamValue("patterns", BinCutFinalInstanceRow.PatSetName);
                function.SetParamValue("performanceMode", (isHbv ? "HBV_" : "") + SourceRow.PerformanceMode);
                function.SetParamValue("isHarvesting", BinCutFinalInstanceRow.BinCutInstanceRow.IsHarvesting);
                if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.BinOutStage))
                {
                    HashSet<string> binOutJobs = BinCutFinalInstanceRow.BinCutInstanceRow.BinOutStage.Split(',').ToHashSet(StringComparer.OrdinalIgnoreCase);
                    string disableBinOutJobs = string.Join(",", AllJobs.Where(x => !binOutJobs.Contains(x)));
                    function.SetParamValue("disableBinOut", binOutJobs.Contains("x") ? "ALL" : disableBinOutJobs);
                }
                if (BinCutFinalInstanceRow.CanBeBurst)
                {
                    function.SetParamValue("resultMode", "1");
                }
            }
            else
            {
                function.SetParamValue("forceE1Modes", SourceRow.ColumnContent.Split(new[] { " ", ":" }, StringSplitOptions.RemoveEmptyEntries).First());
            }
            function.SetParamValue("instanceFailFlag", isHbv ? GetFlag() : GetFlagBv());

            UserFunctionTableRow ufFuncSetting = null;
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.UserFunction) && TestPlanStatic.UserFunctionSheet != null)
            {
                ufFuncSetting = TestPlanStatic.UserFunctionSheet.Rows
                .FirstOrDefault(x => x.UserFunction.Equals(BinCutFinalInstanceRow.BinCutInstanceRow.UserFunction, StringComparison.OrdinalIgnoreCase));
            }
            List<string> ufDigSrcPats = TestPlanStatic.UfDigSrcSheets
                .SelectMany(x => x.Rows)
                .Where(y => !string.IsNullOrEmpty(y.PatternName)).Select(z => z.PatternName).ToList();
            AtpgService.SetDigSrc(
                BinCutFinalInstanceRow,
                ufDigSrcPats,
                ufFuncSetting,
                LocalSpecs.HardIpInfos, $":{(isHbv ? "HBV_" : "") + SourceRow.PerformanceMode}",
                BinCutFinalInstanceRow.PatternList,
                ref function
            );
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;
        }

        private void GenerateVbtInstanceRow(Function function, InstanceRow row, string vbtName)
        {
            string dsscPat = "";
            if (vbtName.StartsWith("GradeSearch_"))
            {
                function.ArgList[0] = BinCutFinalInstanceRow.PatSetName;
                function.SetParamValue("Performance_mode", GetBinningDomain());
                function.SetParamValue("Offset_testType", SourceRow.ColumnName.ToString());
                function.SetParamValue("HarvPinGrp_Enable", BinCutFinalInstanceRow.BinCutInstanceRow.HarvPinGrpEnable);
                function.SetParamValue("HarvestBinningFlag", BinCutFinalInstanceRow.BinCutInstanceRow.HarvestBinningFlag);
                function.SetParamValue("UserFunction", BinCutFinalInstanceRow.BinCutInstanceRow.UserFunction);
                if (Regex.IsMatch(BinCutFinalInstanceRow.BinCutInstanceRow.MultiFstpEnable, "TRUE", RegexOptions.IgnoreCase))
                {
                    function.SetParamValue("MultiFSTP_Enable", "TRUE");
                }

                if (vbtName.Equals("GradeSearch_HVCC_VT", StringComparison.OrdinalIgnoreCase) && BinCutFinalInstanceRow.BinCutInstanceRow.HarvPinGrpEnable.Equals("True", StringComparison.OrdinalIgnoreCase))
                {
                    function.SetParamValue("Harv_CommonFailFlag", GetFlag());
                }

                if (BinCutFinalInstanceRow.CanBeBurst)
                {
                    function.SetParamValue("result_mode", "1");
                    function.SetParamValue("DecomposePatt", "No");
                }
            }
            else
            {
                function.SetParamValue("str_setE1_flag", SourceRow.ColumnContent.Split(new[] { " ", ":" }, StringSplitOptions.RemoveEmptyEntries).Last());
                function.SetParamValue("power_pin", "VDD_" + SourceRow.BinningDomain);
                function.SetParamValue("performance_mode", SourceRow.ColumnContent.Split(new[] { " ", ":" }, StringSplitOptions.RemoveEmptyEntries).First());
            }
            if (BinCutFinalInstanceRow.PatternList.Any())
            {
                foreach (string pat in BinCutFinalInstanceRow.PatternList)
                {
                    if (Regex.IsMatch(pat, @"\w*DSSC\w*", RegexOptions.IgnoreCase))
                    {
                        dsscPat = pat;
                    }
                }
            }
            if (!string.IsNullOrEmpty(dsscPat))
            {
                string sendPinName = "JTAG_TDI";
                if (LocalSpecs.HardIpInfos != null)
                {
                    HardIpInfo target = LocalSpecs.HardIpInfos.GetHardIpInfo(dsscPat);
                    sendPinName = string.IsNullOrEmpty(target.SendPinName) ? sendPinName : target.SendPinName;
                }
                function.SetParamValue("DigSrc_Pin", sendPinName);
            }
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;
        }

        protected string GetDomain()
        {
            if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConCpu + "_").ToUpper()))
            {
                return BinCutConstant.ConCpu;
            }

            if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConGpu + "_").ToUpper()))
            {
                return BinCutConstant.ConGpu;
            }

            if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConSoc + "_").ToUpper()))
            {
                return BinCutConstant.ConSoc;
            }

            if (SourceRow.ColumnName == EnumColumnName.TD)
            {
                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConCpu + "Td").ToUpper()))
                {
                    return BinCutConstant.ConCpu;
                }

                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConGpu + "Td").ToUpper()))
                {
                    return BinCutConstant.ConGpu;
                }

                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConSoc + "Td").ToUpper()))
                {
                    return BinCutConstant.ConSoc;
                }

                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConCpu + "Sa").ToUpper()))
                {
                    return BinCutConstant.ConCpu;
                }

                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConGpu + "Sa").ToUpper()))
                {
                    return BinCutConstant.ConGpu;
                }

                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConSoc + "Sa").ToUpper()))
                {
                    return BinCutConstant.ConSoc;
                }
            }

            if (SourceRow.ColumnName == EnumColumnName.Mbist)
            {
                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConCpu + "Mbist").ToUpper()))
                {
                    return BinCutConstant.ConCpu;
                }

                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConGpu + "Mbist").ToUpper()))
                {
                    return BinCutConstant.ConGpu;
                }

                if (BinCutFinalInstanceRow.PatSetName.ContainsIgnoreCase(("_" + BinCutConstant.ConSoc + "Mbist").ToUpper()))
                {
                    return BinCutConstant.ConSoc;
                }
            }

            return "";
        }

        protected virtual string GetBinningDomain()
        {
            return SourceRow.GetBinningDomain();
        }

        internal string GetDcSelector(string selectorName = null)
        {
            if (!string.IsNullOrEmpty(selectorName))
            {
                if (selectorName.Equals("HV", StringComparison.CurrentCultureIgnoreCase))
                {
                    return Max;
                }

                if (selectorName.Equals("LV", StringComparison.CurrentCultureIgnoreCase))
                {
                    return Min;
                }

                if (selectorName.Equals("NV", StringComparison.CurrentCultureIgnoreCase))
                {
                    return Typ;
                }
            }
            if (Regex.IsMatch(SourceRow.AllOther, @"HV\s*Levels"))
            {
                return Max;
            }

            if (Regex.IsMatch(SourceRow.AllOther, @"NV\s*Levels"))
            {
                return Typ;
            }

            if (Regex.IsMatch(SourceRow.AllOther, @"LV\s*Levels"))
            {
                return Min;
            }

            if (SourceRow.AllOther.EndsWith("HV", StringComparison.CurrentCultureIgnoreCase))
            {
                return Max;
            }

            if (SourceRow.AllOther.EndsWith("NV", StringComparison.CurrentCultureIgnoreCase))
            {
                return Typ;
            }

            return Min;
        }

        protected string GetAcCategory(string timeSet, BlockType type)
        {
            string category = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSet, type);
            if (category == "TBD")
            {
                category = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSet);
            }

            return category;
        }

        protected string GetDcCategory()
        {
            string dcCategory = SourceRow.GetDcCategory();
            if (BinCutFinalInstanceRow.BinCutInstanceRow != null && !string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.DCcategory))
            {
                dcCategory = BinCutFinalInstanceRow.GetDcCategory();
            }

            if (MultiTestSettingSheetsSingleton.Instance() != null)
            {
                if (MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.Exists(s => s.CategoryName.Equals(dcCategory, StringComparison.OrdinalIgnoreCase)))
                {
                    return dcCategory;
                }

                return MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.FindBinCutCategoryName(TestPlanStatic.PowerInfoSheet, "", GetDomain(), out EnumMessageLevel _, out string _);
            }

            return dcCategory;
        }

        internal virtual string GetInstanceName()
        {
            string modePatSetName = GetModePatSetName();
            string parameter = BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Hardip || BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Rtos ? BinCutFinalInstanceRow.BinCutInstanceRow.Type + "_" + modePatSetName : modePatSetName;
            parameter += $"_{BinCutFinalInstanceRow.BinCutInstanceRow.Instance}";
            if (SourceRow.ColumnContent.ContainsIgnoreCase("elb") || SourceRow.ColumnContent.ContainsIgnoreCase("ilb"))
            {
                parameter += "_CallInst";
            }
            parameter = DeleteModeInInstanceName(parameter);
            parameter = AdditionInfoInstanceName(parameter);
            return parameter + "_" + SourceRow.GetBinType();
        }

        protected string GetIlbElbInstanceName()
        {
            string modePatSetName = GetModePatSetName();
            string parameter = BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Hardip || BinCutFinalInstanceRow.BinCutInstanceRow.Type == BincutInstanceType.Rtos ? BinCutFinalInstanceRow.BinCutInstanceRow.Type + "_" + modePatSetName : modePatSetName;
            parameter += $"_{BinCutFinalInstanceRow.BinCutInstanceRow.Instance}";
            parameter += SourceRow.TableType == EnumBinCutTableType.Post ? "_outsidebincut" : "_CallInst";
            parameter = DeleteModeInInstanceName(parameter);
            parameter = AdditionInfoInstanceName(parameter);
            return parameter + "_" + SourceRow.GetBinType();
        }

        protected string GetModePatSetName()
        {
            var instSheetPreProcess = new InstSheetPreProcessBinCut(BinCutInputManager.Config);
            string mode = instSheetPreProcess.GetMode(BinCutFinalInstanceRow.BinCutInstanceRow);
            string patSetName = string.IsNullOrEmpty(BinCutFinalInstanceRow.FinalInstName) ? BinCutFinalInstanceRow.PatSetName : BinCutFinalInstanceRow.FinalInstName;
            return Combination.CombineByUnderLine(new List<string> { mode, patSetName });
        }

        internal string DeleteModeInInstanceName(string para)
        {
            List<string> nameList = para.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (nameList.Count < 2)
            {
                return para;
            }

            if (nameList[1].Equals(SourceRow.TargetPerformanceMode))
            {
                nameList.RemoveAt(1);
                return string.Join("_", nameList);
            }
            return para;
        }

        internal string AdditionInfoInstanceName(string para)
        {
            string keyword = SourceRow.PerformanceMode;
            if (Regex.IsMatch(keyword, "_TMPS", RegexOptions.IgnoreCase))
            {
                keyword = Regex.Replace(keyword, "_TMPS", "", RegexOptions.IgnoreCase);
            }
            return keyword + "_" + para;
        }

        protected virtual string GetVbtName()
        {
            return SourceRow.GetVbtFunction();
        }

        protected virtual string GetVbtNameCs()
        {
            if (!string.IsNullOrEmpty(BinCutFinalInstanceRow.BinCutInstanceRow.FunctionName))
            {
                return BinCutFinalInstanceRow.BinCutInstanceRow.FunctionName;
            }

            return SourceRow.GetVbtFunctionCs();
        }

        internal virtual string GenerateAcSelector()
        {
            return Typ;
        }

        internal virtual string GenerateDcSelector(string selectName = null)
        {
            return GetDcSelector(selectName);
        }

        #region abstract function
        internal abstract string GenerateAcCategory(InstanceRow pRow);

        internal virtual string GenerateDcCategory()
        {
            return GetDcCategory();
        }

        protected abstract string GenerateLevel();
        #endregion

        public string GetColumnA(EnumColumnName enumColumnName, BinCutInstanceRow binCutInstanceRow)
        {
            if (enumColumnName.Equals(EnumColumnName.E1Voltage))
            {
                return "Set E1 voltage instance from BMS";
            }

            if (enumColumnName.Equals(EnumColumnName.RelayOn) || enumColumnName.Equals(EnumColumnName.RelayOff))
            {
                return "Relay instance from BMS";
            }

            if (enumColumnName.Equals(EnumColumnName.CallNwireEnable) || enumColumnName.Equals(EnumColumnName.CallNwireDisable))
            {
                return "Call NwireFlow from BMS";
            }

            if (enumColumnName.Equals(EnumColumnName.CallTMPS))
            {
                return "Call TMPS from BMS";
            }

            return string.IsNullOrEmpty(binCutInstanceRow.FlowName)
                ? binCutInstanceRow.SheetName + ",Row" + binCutInstanceRow.RowNum
                : binCutInstanceRow.SheetName + ",Row" + binCutInstanceRow.RowNum + "," + binCutInstanceRow.FlowName + "," + enumColumnName;
        }
    }
}
