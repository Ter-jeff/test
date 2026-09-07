using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.PostAction.TempMon.Enums;
using Automation.GenerateIgxl.Scan.Harvest;
using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Atpg;
using Automation.Utility.Basic;
using Automation.Utility.HardIP;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Concurrent;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HTOL
{
    public class HtolInstanceMain : ScanNonBinCutInstanceMain
    {
        protected override string Module => "HTOL";
        private readonly List<BinCutInstanceSheet> _binCutInstanceSheets;
        private const string BinTableNonBinCut = "Bin_Table_Htol";
        private readonly MultiTestSettingSheetsSingleton _multiTestSettingSheetsSingleton;
        private readonly List<string> _performanceModeList;
        private readonly List<string> _usedInstance = new List<string>();

        public HtolInstanceMain(ScanConfig scanConfig, List<BinCutInstanceSheet> binCutInstanceSheets) : base(scanConfig)
        {
            _binCutInstanceSheets = binCutInstanceSheets;
            _multiTestSettingSheetsSingleton = MultiTestSettingSheetsSingleton.Instance();
            _performanceModeList = MultiTestSettingSheetsSingleton.Instance().PerformanceModeList;
        }

        public override void WorkFlow()
        {
            var patSets = new List<PatSet>();
            var flows = new List<SubFlowSheet>();
            var instances = new List<InstanceRow>();
            var instancesProfile = new List<InstanceRow>();
            var binTableRows = new List<BinTableRow>();
            var binTableSheet = new BinTableSheet(BinTableNonBinCut);

            BinCutInstanceNamingSheet binCutInstanceNamingSheet = BinCutInstanceNamingSheet();

            var instSheetPreProcess = new InstSheetPreProcess(Config);
            BinCutFinalInstanceRows binCutFinalInstanceRows = instSheetPreProcess.InitialInstance(_binCutInstanceSheets, binCutInstanceNamingSheet);
            binCutFinalInstanceRows = binCutFinalInstanceRows.RePatSetNameDuplicateRows();

            AcSpecSheet acSpecSheet = TestProgram.IgxlWorkBk.GetAcSpecsSheet();
            if (acSpecSheet != null)
            {
                new BinCutAcSpecsWriter().GenAcSpecs(binCutFinalInstanceRows, acSpecSheet);
            }

            instances.AddRange(GenInstances(binCutFinalInstanceRows, ref instancesProfile));

            flows.AddRange(GenFlow(binCutFinalInstanceRows));

            binTableRows.AddRange(GetBinTableRows(binCutFinalInstanceRows));
            List<PatSet> patSetSheet = GenPatSets(binCutFinalInstanceRows);
            patSetSheet = AddCommandAndFlagInPatSet(patSetSheet, binCutFinalInstanceRows);
            patSets.AddRange(patSetSheet);

            #region Add into igxl
            var list = binTableRows.GroupBy(x => x.Name).ToList();
            foreach (IGrouping<string, BinTableRow> grouping in list)
            {
                binTableSheet.AddRow(grouping.First());
            }

            TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirNonBinCut, binTableSheet);

            List<string> flags = GenAllFailFlagToMainFlow(binTableSheet);
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(flags, "Scan", FolderStructure.DirMain);

            SetPatSetSheet(patSets);

            foreach (SubFlowSheet flow in flows)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirNonBinCut, flow);
            }

            SetInstanceSheet(instances, instancesProfile);
            #endregion
            AddControlBinFlagsToMainInitEnableWd();
        }

        internal override List<PatSet> GenPatSets(List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            var patSets = new List<PatSet>();
            foreach (BinCutFinalInstanceRow binCutFinalInstanceRow in binCutFinalInstanceRows)
            {
                if (binCutFinalInstanceRow.PatternList.Count < 2)
                {
                    continue;
                }

                var patSet = new PatSet { PatSetName = binCutFinalInstanceRow.PatSetName };

                foreach (string pattern in binCutFinalInstanceRow.PatternList)
                {
                    var patSetRow = new PatSetRow
                    {
                        File = pattern,
                        Burst = binCutFinalInstanceRow.IsBurstNonBinCutInstance() ? "Yes" : "No",
                        Comment = "SheetName: " + binCutFinalInstanceRow.BinCutInstanceRow.SheetName + ", Row: " + binCutFinalInstanceRow.BinCutInstanceRow.RowNum
                    };

                    patSet.AddRow(patSetRow);
                }
                if (!patSets.Exists(x => x.PatSetName.Equals(patSet.PatSetName)))
                {
                    patSets.Add(patSet);
                }
            }
            return patSets;
        }

        internal override void SetPatSetSheet(List<PatSet> patSets)
        {
            if (patSets.Count > 0)
            {
                var patSetSheet = new PatSetSheet("PatSets_Htol");
                patSetSheet.AddRows(patSets);
                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirNonBinCut, patSetSheet);
            }
        }

        private void SetInstanceSheet(List<InstanceRow> instanceRows, List<InstanceRow> instancesProfile)
        {
            if (instanceRows.Count > 0)
            {
                var instanceSheet = new InstanceSheet("TestInst_Htol");
                instanceSheet.AddRows(instanceRows);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirNonBinCut, instanceSheet);

                var instProfileSheet = new InstanceSheet("TestInst_Profile");
                instProfileSheet.AddRows(instancesProfile);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirNonBinCut, instProfileSheet);
            }
        }

        internal List<InstanceRow> GenInstances(List<BinCutFinalInstanceRow> binCutFinalInstanceRows, ref List<InstanceRow> profiles)
        {
            var instanceRows = new InstanceRows();
            var concurrentFlow = new ConcurrentFlowSheet("Concurrent_Flow");
            if (EpWorkbook.TestPlanWorkbook.Worksheets["Concurrent_Flow"] != null)
            {
                concurrentFlow = TestPlanStatic.ConcurrentFlowSheet;
            }
            IEnumerable<IGrouping<string, BinCutFinalInstanceRow>> groups = binCutFinalInstanceRows.GroupBy(x => x.BinCutInstanceRow.FlowName.ToUpper());
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                foreach (BinCutFinalInstanceRow row in group)
                {
                    var instanceRow = new InstanceRow();
                    if (IsAtpg(row))
                    {
                        instanceRow = CreateAtpgInstance(row, instanceRows, concurrentFlow);
                        instanceRows.Add(instanceRow);
                    }
                    else
                    {
                        foreach (string pat in row.PatternList)
                        {
                            instanceRow = CreateHardIpInstance(row, instanceRows, pat);
                            if (!_usedInstance.Contains(instanceRow.TestName))
                            {
                                _usedInstance.Add(instanceRow.TestName);
                                instanceRows.Add(instanceRow);
                            }
                        }
                    }
                    TempMonService.TrySetTempMon(LocalSpecs.TempMonDatas, row.BinCutInstanceRow.TempMon, instanceRow.TestName, EnumType.Instance);
                    profiles.AddRange(GenProfileInstances(instanceRow.TestName));
                }

                instanceRows.AddHeaderFooter(group.Key);
            }
            return instanceRows;
        }

        protected override string GetPerformanceModes(BinCutFinalInstanceRow row)
        {
            string mode = string.IsNullOrEmpty(row.PerformanceMode) ? AtpgService.GetPerformanceMode(row.InitList, _performanceModeList) : row.PerformanceMode;
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

        internal override string GetDcCategory(BinCutFinalInstanceRow binCutFinalInstanceRow)
        {
            string mode = GetPerformanceModes(binCutFinalInstanceRow);
            if (!string.IsNullOrEmpty(binCutFinalInstanceRow.BinCutInstanceRow.DCcategory))
            {
                string dcCategoryUserDefine = GetDcCategory(binCutFinalInstanceRow.BinCutInstanceRow.DCcategory);
                if (dcCategoryUserDefine.Contains("_EQN"))
                {
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
                dcCategory = "HTOL_X_X_X";
            }

            return dcCategory;
        }

        private InstanceRow CreateAtpgInstance(BinCutFinalInstanceRow row, List<InstanceRow> rows, ConcurrentFlowSheet concurrentFlow)
        {
            var instanceRow = new InstanceRow();
            string selector = row.GetVoltageType();
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
            }

            instanceRow.TestName = row.GetParameter();
            if (rows.Exists(x => x.TestName.Equals(instanceRow.TestName, StringComparison.CurrentCultureIgnoreCase)))
            {
                row.IsDuplicateName = true;
                instanceRow.TestName = row.GetParameter();
            }
            instanceRow.TimeSets = row.GetTimeSetVersion(row.PatternList);
            instanceRow.DcCategory = GetDcCategory(row);
            instanceRow.DcSelector = GetDcSelector(selector);
            instanceRow.AcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow?.AcSpec) ? GetAcCategory(row.BinCutInstanceRow, instanceRow.TimeSets) : row.BinCutInstanceRow.AcSpec;
            instanceRow.AcSelector = "Typ";
            instanceRow.PinLevels = concurrentFlow.Rows.Exists(x => x.Subflows.Exists(y => y.ToUpper().Equals(row.BinCutInstanceRow.FlowName.ToUpper())))
                ? GenerateLevelConcurrent(row.BinCutInstanceRow?.FlowName, instanceRow.TimeSets, concurrentFlow, row.BinCutInstanceRow.Levels)
                    : GenerateLevel(instanceRow.DcCategory, row.Block, row.BinCutInstanceRow.Levels);
            if (row.BinCutInstanceRow != null)
            {
                instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
            }

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameFunctionalT, "htol", true);
            if (function.Type == ".NET")
            {
                return GenerateCSharpInstanceRow(row, instanceRow, function);
            }

            function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameFunctionalT, "htol");
            return GenerateVbtInstanceRow(row, instanceRow, function);
        }

        private InstanceRow GenerateCSharpInstanceRow(BinCutFinalInstanceRow row, InstanceRow instanceRow, Function function)
        {
            if (row.GetDcCategory().Contains(instanceRow.DcCategory + "_EQN") && LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                function.SetParamValue("ateTestCondition", row.BinCutInstanceRow.DCcategory);
                List<string> items = row.PatSetName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                items.Add(row.GetDcCategory().Replace(instanceRow.DcCategory + "_", ""));
                row.PatSetName = string.Join("_", items.Where(x => !string.IsNullOrEmpty(x)));
                instanceRow.TestName = row.GetParameter();
            }
            function.SetParamValue("Patterns", row.PatternList.Count == 1 ? row.PatternList[0] : row.PatSetName);
            function.SetParamValue("ResultMode", row.IsBurstNonBinCutInstance() ? "1" : "0");
            function.SetParamValue("RelayMode", "1");

            function.SetParamValue("harvestOtherFailFlag", string.Join(",", row.GetPinGroupFailFlags()));
            function.SetParamValue("isHarvesting", row.BinCutInstanceRow?.IsHarvesting);

            UserFunctionTableRow ufFuncSetting = null;
            if (!string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction) && TestPlanStatic.UserFunctionSheet != null)
            {
                ufFuncSetting = TestPlanStatic.UserFunctionSheet.Rows
                .FirstOrDefault(x => x.UserFunction.Equals(row.BinCutInstanceRow.UserFunction, StringComparison.OrdinalIgnoreCase));
            }
            if (ufFuncSetting != null)
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, function);
            }
            List<string> ufDigSrcPats = TestPlanStatic.UfDigSrcSheets
                .SelectMany(x => x.Rows)
                .Where(y => !string.IsNullOrEmpty(y.PatternName)).Select(z => z.PatternName).ToList();
            AtpgService.SetDigSrc(row, ufDigSrcPats, ufFuncSetting, LocalSpecs.HardIpInfos, "", row.PatternList, ref function);
            instanceRow.VbtType = ".NET";
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        private InstanceRow GenerateVbtInstanceRow(BinCutFinalInstanceRow row, InstanceRow instanceRow, Function function)
        {
            if (row.GetDcCategory().Contains(instanceRow.DcCategory + "_EQN") && LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                function.SetParamValue("ApplyVoltageFromBinCut", row.BinCutInstanceRow.DCcategory);
                List<string> items = row.PatSetName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                items.Add(row.GetDcCategory().Replace(instanceRow.DcCategory + "_", ""));
                row.PatSetName = string.Join("_", items.Where(x => !string.IsNullOrEmpty(x)));
                instanceRow.TestName = row.GetParameter();
            }
            function.SetParamValue("Patterns", row.PatternList.Count == 1 ? row.PatternList[0] : row.PatSetName);
            function.SetParamValue("ResultMode", row.IsBurstNonBinCutInstance() ? "1" : "0");
            function.SetParamValue("RelayMode", "1");

            if (Regex.IsMatch(row.BinCutInstanceRow.MultiFstpEnable, "TRUE", RegexOptions.IgnoreCase))
            {
                function.SetParamValue("MultiFSTP_Enable", "TRUE");
            }

            instanceRow.VbtType = "VBT";
            instanceRow.VbtName = function.FunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        private InstanceRow CreateHardIpInstance(BinCutFinalInstanceRow row, List<InstanceRow> rows, string pattern)
        {
            var instanceRow = new InstanceRow();
            string selector = row.GetVoltageType();
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
            }

            instanceRow.TestName = $"HIP_{pattern.ToUpper()}";
            if (rows.Exists(x => x.TestName.Equals(instanceRow.TestName, StringComparison.CurrentCultureIgnoreCase)))
            {
                row.IsDuplicateName = true;
                instanceRow.TestName = $"HIP_{pattern.ToUpper()}";
            }
            instanceRow.TimeSets = row.GetTimeSetVersion(row.PatternList);
            instanceRow.DcCategory = GetDcCategory(row);
            instanceRow.DcSelector = GetDcSelector(selector);
            instanceRow.AcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow?.AcSpec) ? GetAcCategory(row.BinCutInstanceRow, instanceRow.TimeSets) : row.BinCutInstanceRow.AcSpec;
            instanceRow.AcSelector = "Typ";
            instanceRow.PinLevels = "Levels_HardIP";

            if (row.BinCutInstanceRow != null)
            {
                instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
            }

            Function vbtFunctionBase;
            HardIpInfo info = HardIpService.GetHardIpInfo(pattern);
            if (info.CapBit > 0)
            {
                string userFunction = "";
                if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction))
                {
                    userFunction = row.BinCutInstanceRow.UserFunction;
                }

                vbtFunctionBase = SetHardIpParameter(info, pattern, userFunction);
            }
            else
            {
                vbtFunctionBase = SetFunctionTParameter(row, pattern, instanceRow);
            }

            instanceRow.VbtName = vbtFunctionBase.FullFunctionName;
            instanceRow.VbtType = vbtFunctionBase.Type;
            instanceRow.ArgList = vbtFunctionBase.Parameters;
            instanceRow.Args = vbtFunctionBase.ArgList;

            return instanceRow;
        }

        private bool IsAtpg(BinCutFinalInstanceRow row)
        {
            foreach (string pattern in row.PatternList)
            {
                string[] arr = pattern.Split('_');
                if (arr.Length > 4)
                {
                    string type = arr[4].ToUpper();
                    if (type.Contains("SC") || type.Contains("BI"))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private Function SetFunctionTParameter(BinCutFinalInstanceRow row, string pattern, InstanceRow instance)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameFunctionalT, "htol", true);

            if (function.Type == ".NET")
            {
                instance.VbtType = ".NET";
                function.SetParamValue("Patterns", pattern);
                function.SetParamValue("ResultMode", row.IsBurstNonBinCutInstance() ? "1" : "0");

                function.SetParamValue("RelayMode", "1");

                if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction))
                {
                    TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, function);
                }
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameFunctionalT, "htol");
                instance.VbtType = "VBT";
                function.SetParamValue("Patterns", pattern);
                function.SetParamValue("ResultMode", row.IsBurstNonBinCutInstance() ? "1" : "0");

                function.SetParamValue("RelayMode", "1");

                if (Regex.IsMatch(row.BinCutInstanceRow.MultiFstpEnable, "TRUE", RegexOptions.IgnoreCase))
                {
                    function.SetParamValue("MultiFSTP_Enable", "TRUE");
                }
            }
            return function;
        }

        private Function SetHardIpParameter(HardIpInfo info, string pattern, string userFunction)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(VbtFunctionLibShared.VifName, "htol");

            function.SetParamValue(function.Type == ".NET" ? "patternSet" : "patset", pattern);
            function.SetParamValue(function.Type == ".NET" ? "digCapPin" : "DigCap_Pin", info.CapPinName);
            //DigCap_DataWidth:  Get from "Cap Bit Str" in patInfo file. Like "wdr14_10+wdr23_10" ===> 10
            //function.ArgList[18] = Regex.Match(info.CapBitStr, @"^wdr\d+_(?<num>(\d+)).*").Groups["num"].ToString();
            function.SetParamValue(function.Type == ".NET" ? "digCapDataWidth" : "DigCap_DataWidth", SearchInfo.GetDigDataWidth(info.CapBitStr));
            //DigCap_Sample_Size: Get from "Cap Bit" in patInfo file
            //function.ArgList[19] = info.CapBit.ToString();
            function.SetParamValue(function.Type == ".NET" ? "digCapSampleSize" : "DigCap_Sample_Size", info.CapBit.ToString("G15"));

            //CUS_Str_DigCapData
            //function.ArgList[31] = SearchInfo.GetCusStrDigCapData(pattern);
            function.SetParamValue(function.Type == ".NET" ? "digCapDataCustomString" : "CUS_Str_DigCapData", info.DsscOut);

            if (!string.IsNullOrEmpty(userFunction))
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(userFunction, function);
            }
            return function;
        }

        private List<InstanceRow> GenProfileInstances(string instName)
        {
            var result = new List<InstanceRow>();
            var instStart = new InstanceRow { TestName = instName + "_Current_Profile_Start" };
            Function functionProfileStart = TestProgram.VbtFunctionLib.GetFunctionByName("ProfileAutoResolution", "htol", true);
            if (functionProfileStart.Type != ".NET")
            {
                functionProfileStart = TestProgram.VbtFunctionLib.GetFunctionByName("Start_Profile_AutoResolution", "htol");
            }

            functionProfileStart.SetParamValue("PinName", "CorePower");
            functionProfileStart.SetParamValue("WhatToCapture", "I");
            functionProfileStart.SetParamValue("CapSignalName", "Capture_Signal");
            functionProfileStart.SetParamValue("Plottime", "10");
            instStart.VbtType = functionProfileStart.Type;
            instStart.VbtName = functionProfileStart.FullFunctionName;
            instStart.ArgList = functionProfileStart.Parameters;
            instStart.Args = functionProfileStart.ArgList;
            result.Add(instStart);

            var instanceRow = new InstanceRow { TestName = instName + "_Profile_Plot" };
            Function functionProfilePlot = TestProgram.VbtFunctionLib.GetFunctionByName("PlotProfile", "htol", true);
            if (functionProfilePlot.Type != ".NET")
            {
                functionProfilePlot = TestProgram.VbtFunctionLib.GetFunctionByName("Plot_Profile", "htol");
            }

            functionProfilePlot.SetParamValue("PinName", "CorePower");
            functionProfilePlot.SetParamValue("CapSignalName", "Capture_Signal");
            functionProfilePlot.SetParamValue("ExportWaveform", "TRUE");
            functionProfilePlot.SetParamValue("PlotWaveform", "FALSE");
            functionProfilePlot.SetParamValue("Calculate_ProfileInfo", "FALSE");
            instanceRow.VbtType = functionProfilePlot.Type;
            instanceRow.VbtName = functionProfilePlot.FullFunctionName;
            instanceRow.ArgList = functionProfilePlot.Parameters;
            instanceRow.Args = functionProfilePlot.ArgList;
            result.Add(instanceRow);
            return result;
        }

        protected override List<FlowRow> GenFlowBody(IGrouping<string, BinCutFinalInstanceRow> group)
        {
            var flowRows = new List<FlowRow>();
            foreach (BinCutFinalInstanceRow dataRow in group)
            {
                var rows = new List<FlowRow>();
                FlowRow testFlowRow = null;
                if (IsAtpg(dataRow))
                {
                    testFlowRow = GetTestRow(dataRow);
                    rows.Add(testFlowRow);
                    if (!dataRow.BinCutInstanceRow.IsBypassBinOut)
                    {
                        rows.Add(GetControlBinRow(dataRow));
                    }
                }
                else
                {
                    foreach (string pat in dataRow.PatternList)
                    {
                        testFlowRow = GetTestRow(dataRow, $"HIP_{pat.ToUpper()}");
                        rows.Add(testFlowRow);
                        if (!dataRow.BinCutInstanceRow.IsBypassBinOut)
                        {
                            rows.Add(GetControlBinRow(dataRow, $"HIP_{pat.ToUpper()}"));
                        }
                        HardIpInfo info = HardIpService.GetHardIpInfo(pat);
                        if (!string.IsNullOrEmpty(info.CapBitName))
                        {
                            foreach (string reg in info.CapBitName.Split('+'))
                            {
                                var rowTmp = new FlowRow
                                {
                                    Opcode = "use_limit",
                                    Parameter = testFlowRow.Parameter,
                                    TName = reg
                                };
                                rows.Add(rowTmp);
                            }
                        }
                    }
                }

                var startRow = new FlowRow
                {
                    Parameter = testFlowRow?.Parameter + "_Current_Profile_Start",
                    Opcode = "Test",
                    ColumnA = testFlowRow?.ColumnA,
                    Enable = "CurrentProfile"
                };

                flowRows.Add(startRow);

                flowRows.Add(new FlowRow { Opcode = "loop", Parameter = "10" });
                flowRows.AddRange(rows);
                flowRows.Add(new FlowRow { Opcode = "endloop" });

                var flowRow = new FlowRow
                {
                    Parameter = testFlowRow?.Parameter + "_Profile_Plot",
                    Opcode = "Test",
                    ColumnA = testFlowRow?.ColumnA,
                    Enable = "CurrentProfile||VoltageProfile"
                };
                flowRows.Add(flowRow);
            }
            return flowRows;
        }
    }
}
