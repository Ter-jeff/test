using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.PostAction.TempMon.Enums;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;
using Automation.Utility.Basic;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinNumber;
using TestPlanLib.EVS;
using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.EVS
{
    public class EvsMultiRampInstanceMain : EvsInstanceMainCs
    {
        private const string EvsTightFlag = "F_EVS_Tight";
        private const string EvsPwrDownFLag = "F_EVS_PwrDown";
        private readonly Dictionary<string, string> _evsFailFlag = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public EvsMultiRampInstanceMain(ScanConfig config, List<BinCutInstanceSheet> evsInstanceSheets) : base(config, evsInstanceSheets)
        {
        }

        public List<InstanceRow> GenRampEvsInstances(IEnumerable<BinCutFinalInstanceRow> tpInstRows)
        {
            var allInstanceRows = new InstanceRows();
            foreach (BinCutFinalInstanceRow row in tpInstRows)
            {
                var instanceRows = new InstanceRows();
                //Generate pattern instance
                InstanceRow normalInstanceRow = GenEvsNormalInstance(row);
                if (!instanceRows.Exists(x => x.TestName.Equals(normalInstanceRow.TestName)))
                {
                    instanceRows.Add(normalInstanceRow);
                }
                //Generate ramp instances
                foreach (EvsCondition condition in row.BinCutInstanceRow.Evs.EvsConditions)
                {
                    if (!string.IsNullOrEmpty(condition.Voltage1))
                    {
                        string instancePinList;
                        Function function;
                        InstanceRow instanceRow;
                        (instanceRow, instancePinList, function) = GenEvsRampInstance(row, condition);
                        instanceRows.Add(instanceRow);
                        instanceRows.Add(GenEvsVTrigInstance(instanceRow, function, condition));
                        instanceRows.Add(GenEvsIvCurveInstance(instanceRow, function, condition));
                        instanceRows.Add(GenEvsCurrentProfileStartInstance(instanceRow, row, instancePinList, condition));
                        instanceRows.Add(GenEvsCurrentProfilePlotInstance(instanceRow, row, instancePinList, condition));
                    }
                }
                instanceRows.ForEach(x => TempMonService.TrySetTempMon(LocalSpecs.TempMonDatas, row.BinCutInstanceRow.TempMon, x.TestName, EnumType.Instance));
                allInstanceRows.AddRange(instanceRows);
            }
            return allInstanceRows;
        }

        private FlowRow GenEvsIvCurveTestRow(FlowRow flowRow)
        {
            FlowRow copy = flowRow.Copy();
            copy.ColumnA = "IV curve for " + flowRow.ColumnA;
            copy.Parameter = flowRow.Parameter + "_IV";
            copy.Enable = "IVCurve";
            copy.FailAction = EvsTightFlag;

            return copy;
        }

        private FlowRow GenEvsVTrigTestRow(FlowRow flowRow)
        {
            FlowRow copy = flowRow.Copy();
            copy.ColumnA = "Vtrig for " + flowRow.ColumnA;
            copy.Parameter = flowRow.Parameter + "_Vtrig";
            copy.Enable = "Vtrig";
            copy.FailAction = EvsTightFlag;

            return copy;
        }

        private List<FlowRow> GetEvsTestRow(BinCutFinalInstanceRow row, string failFlag, EvsCondition condition, bool isRamp)
        {
            var flowRow = new FlowRow();
            var flowRows = new List<FlowRow>();
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                flowRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
            }

            flowRow.Opcode = row.NopByEnableWord ? OpCode.Nop : OpCode.Test;
            flowRow.Parameter = isRamp ? GetEvsRampPowerName(row, condition) : GetEvsNormalInstanceName(row);

            flowRow.Job = condition.JobStage;
            flowRow.Enable = row.GetEnable();
            if (isRamp)
            {
                flowRow.Enable = string.IsNullOrEmpty(flowRow.Enable) ? "A_EVS" : flowRow.Enable + "&& A_EVS";
            }

            flowRow.FailAction = string.IsNullOrEmpty(failFlag) ? "" :
                isRamp && condition.EvsNum.Equals("EVS1", StringComparison.CurrentCultureIgnoreCase) ? string.Join(",", new string[] { failFlag, EvsTightFlag, EvsPwrDownFLag }) : failFlag;
            flowRows.Add(flowRow);

            if (isRamp)
            {
                if (condition.UserFunction.TryGetValue("Ramp", out string valueRamp) && !string.IsNullOrEmpty(valueRamp))
                {
                    flowRows.AddRange(TestPlanStatic.UserFunctionSheet.CallAfterInstance(valueRamp));
                }

                if (condition.EvsNum.Equals("EVS1"))
                {
                    flowRows.Add(GenEvsVTrigTestRow(flowRow));
                    if (condition.UserFunction.TryGetValue("Vtrig", out string valueVtrig) && !string.IsNullOrEmpty(valueVtrig))
                    {
                        flowRows.AddRange(TestPlanStatic.UserFunctionSheet.CallAfterInstance(valueVtrig));
                    }

                    flowRows.Add(GenEvsIvCurveTestRow(flowRow));
                    if (condition.UserFunction.TryGetValue("IV", out string valueIV) && !string.IsNullOrEmpty(valueIV))
                    {
                        flowRows.AddRange(TestPlanStatic.UserFunctionSheet.CallAfterInstance(valueIV));
                    }

                    flowRows.Add(FlowRow.GenEndIf(condition.JobStage));
                    flowRows.AddRange(GenPowerResetRows(flowRow));
                    flowRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = EvsPwrDownFLag });
                    flowRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Enable = "Vtrig||IVCurve", Parameter = EvsTightFlag });
                    flowRows.Add(new FlowRow { Opcode = "Next" });
                }
                else
                {
                    flowRows.Add(FlowRow.GenEndIf(condition.JobStage));
                }
            }

            if (isRamp)
            {
                flowRows.Add(new FlowRow { Opcode = "Test", Enable = "CurrentProfile", Parameter = GetProfileInstanceName(row, condition, "Plot") });
            }

            return flowRows;
        }

        public BinTableSheet GenEvsBinTable()
        {
            var binTableSheet = new BinTableSheet("Bin_Table_EVS");
            foreach (KeyValuePair<string, string> flag in _evsFailFlag)
            {
                string domain = flag.Value.Split(',')[0];
                string block = flag.Value.Split(',')[1];
                string bName = flag.Key.Replace("F_EVS", "Bin_EVS");
                var bin = new BinTableRow { Name = bName, ItemList = flag.Key, Op = "OR" };
                bin.Items.Add("T");
                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("EVS", domain, block, bin);
                bin.Sort = binNumInfo.SoftBin.ToString("G15");
                bin.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                bin.Result = binNumInfo.BinNumInfo.Status;

                binTableSheet.AddRow(bin);
            }

            string bNameNoStress = "Bin_EVS_No_Stress";
            var binNoStress = new BinTableRow { Name = bNameNoStress, ItemList = "F_EVS_No_Stress", Op = "OR" };
            binNoStress.Items.Add("T");
            BinNumResult binNumInfoNoStress = BinNumberSingleton.Instance.GetBinInfo("EVS", "No_Stress", "", binNoStress);
            binNoStress.Sort = binNumInfoNoStress.SoftBin.ToString("G15");
            binNoStress.Bin = binNumInfoNoStress.BinNumInfo.HardBin.ToString("G15");
            binNoStress.Result = binNumInfoNoStress.BinNumInfo.Status;

            binTableSheet.AddRow(binNoStress);
            return binTableSheet;
        }

        public List<SubFlowSheet> GenEvsFlow(List<BinCutFinalInstanceRow> tpInstRows)
        {
            var flowSheets = new List<SubFlowSheet>();
            var flowSheetsDic = tpInstRows.GroupBy(x => x.BinCutInstanceRow.SheetName).ToDictionary(y => y.Key.Replace("Instance", "Flow"), y => y.ToList());
            foreach (KeyValuePair<string, List<BinCutFinalInstanceRow>> flow in flowSheetsDic)
            {
                string sourceSheetName = flow.Key.Replace("Flow", "Instance");
                string flowSheetName = flow.Key;
                var flowSequenceList = new List<string>();
                foreach (BinCutFinalInstanceRow row in flow.Value)
                {
                    bool isInstance = !row.BinCutInstanceRow.PatternList.Any() && !row.BinCutInstanceRow.Evs.EvsConditions.Any(x => !string.IsNullOrEmpty(x.Voltage1));
                    if (!flowSequenceList.Exists(x => x.EndsWith(row.BinCutInstanceRow.SubFlow)) && !isInstance)
                    {
                        flowSequenceList.Add(OpCode.Call + " " + row.BinCutInstanceRow.SubFlow);
                    }
                    else if (isInstance)
                    {
                        flowSequenceList.Add(OpCode.Test + " " + row.BinCutInstanceRow.Instance);
                    }
                }
                IEnumerable<IGrouping<string, BinCutFinalInstanceRow>> groupBySubflowName = flow.Value.GroupBy(x => x.BinCutInstanceRow.SubFlow);
                var flowSheet = new SubFlowSheet(flowSheetName, sourceSheetName);
                flowSheet.AddRow(new FlowRow { Opcode = "assign-integer", Parameter = "EVS_Max_Step 1" });
                flowSheet.AddRow(new FlowRow { Opcode = "assign-integer", Enable = "Vtrig", Parameter = "EVS_Max_Step 11" });
                flowSheet.AddRow(new FlowRow { Opcode = "flag-true", Enable = "Vtrig || IVCurve", Parameter = "F_Debug_all" });
                flowSheet.AddRow(new FlowRow { Opcode = "flag-false", Parameter = EvsTightFlag });
                flowSheet.AddRow(new FlowRow { Opcode = "flag-false", Parameter = EvsPwrDownFLag });
                flowSheet.AddRow(new FlowRow { Opcode = "flag-false", Parameter = "F_EVS_No_Stress" });
                if (TestPlanStatic.MainFlowSheet.EvsDeferSubFlowSheets.Any())
                {
                    flowSheet.AddRow(new FlowRow { Opcode = "flag-false", Parameter = "F_EVS_Defer" });
                }

                foreach (IGrouping<string, BinCutFinalInstanceRow> group in groupBySubflowName)
                {
                    if (string.IsNullOrEmpty(group.Key) || group.Key.Equals(flow.Key, StringComparison.CurrentCultureIgnoreCase))
                    {
                        continue;
                    }

                    bool isScan = group.Key.ContainsIgnoreCase("SCAN");
                    SubFlowSheet flowRows = GenEvsSubFlow(group.ToList(), isScan);
                    if (flowRows.Rows.Any())
                    {
                        flowRows.AddEndRows();
                        flowSheets.Add(flowRows);
                    }
                }
                foreach (string flowItem in flowSequenceList)
                {
                    string opcode = flowItem.Split(' ')[0];
                    string parameter = flowItem.Split(' ')[1];
                    flowSheet.AddRow(new FlowRow { Opcode = opcode, Parameter = parameter });
                }

                if (flowSheet.Rows.Any())
                {
                    flowSheet.AddReturnRow();
                    flowSheets.Add(flowSheet);
                }
            }
            return flowSheets;
        }

        private SubFlowSheet GenEvsSubFlow(List<BinCutFinalInstanceRow> instRows, bool isScan)
        {
            var flow = new SubFlowSheet(instRows.First().BinCutInstanceRow.SubFlow);
            if (!instRows.Any())
            {
                return flow;
            }

            var groups = instRows.GroupBy(x => x.Domain).ToList();
            flow.AddStartRows();

            var deferFlags = new List<string>();
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                var binTables = new List<string>();
                var binTableEw = new Dictionary<string, string>();
                for (int i = 0; i < group.Count(); i++)
                {
                    BinCutFinalInstanceRow instRow = group.ElementAt(i);
                    if (instRow.BinCutInstanceRow == null)
                    {
                        continue;
                    }

                    string functionalFailFlag = instRow.BinCutInstanceRow.FailFlag != "" ? instRow.BinCutInstanceRow.FailFlag : "F_EVS_PatSets_" + group.Key + (isScan ? "_Sa_Fail" : "_Bist_Fail");
                    if (!instRow.BinCutInstanceRow.IsBypassBinOut)
                    {
                        AddEvsBinOut(binTables, functionalFailFlag, instRow, binTableEw);
                    }

                    var jobGroups = instRow.BinCutInstanceRow.Evs.EvsConditions.Where(x => !string.IsNullOrEmpty(x.Voltage1)).GroupBy(x => x.JobStage).ToList();
                    foreach (IGrouping<string, EvsCondition> jobGroup in jobGroups)
                    {
                        for (int j = 0; j < jobGroup.Count(); j++)
                        {
                            bool isLastEvsCondition = j == jobGroup.Count() - 1;
                            var multiEvsInstances = new List<FlowRow>();
                            var flowRow = new FlowRow();
                            EvsCondition condition = jobGroup.ElementAt(j);

                            string alarmFailFlag = condition.AlarmFlag != "" ? condition.AlarmFlag : "F_EVS_" + group.Key + (isScan ? "_Sa_Alarm" : "_Bist_Alarm");
                            if (isLastEvsCondition)
                            {
                                AddEvsBinOut(binTables, alarmFailFlag, instRow, binTableEw);
                                deferFlags.Add(alarmFailFlag);
                            }

                            multiEvsInstances.Add(new FlowRow { Opcode = "Test", Enable = "CurrentProfile", Parameter = GetProfileInstanceName(instRow, condition, "Start") });

                            if (!condition.EvsNum.Equals("EVS2"))
                            {
                                multiEvsInstances.Add(new FlowRow { Opcode = "For", Parameter = "EVS_INDEX=0; EVS_INDEX< EVS_Max_Step; EVS_INDEX++" });
                            }
                            multiEvsInstances.Add(condition.EvsNum.Equals("EVS1") ? FlowRow.GenIfCondition("!" + EvsTightFlag, condition.JobStage) : FlowRow.GenIfCondition(EvsTightFlag, condition.JobStage));
                            //Normal instance
                            multiEvsInstances.AddRange(GetEvsTestRow(instRow, functionalFailFlag, condition, false));
                            //Ramp instance
                            multiEvsInstances.AddRange(GetEvsTestRow(instRow, alarmFailFlag, condition, true));

                            flow.AddRows(multiEvsInstances);
                        }
                    }
                }
                foreach (string binTable in binTables)
                {
                    binTableEw.TryGetValue(binTable, out string ewVal);
                    flow.AddRow(new FlowRow { Opcode = OpCode.BinTable, Parameter = binTable, Enable = ewVal });
                }
            }
            if (TestPlanStatic.MainFlowSheet.EvsDeferSubFlowSheets.Any())
            {
                AddJudgeDeferRows(flow, deferFlags);
            }
            return flow;
        }

        private void AddEvsBinOut(List<string> binTables, string failFlag, BinCutFinalInstanceRow instRow, Dictionary<string, string> binTableEw)
        {
            if (!_evsFailFlag.ContainsKey(failFlag))
            {
                _evsFailFlag.Add(failFlag, $"{instRow.Domain},{instRow.Block}");
            }

            if (!binTables.Exists(x => x.Equals(failFlag.Replace("F_EVS", "Bin_EVS"), StringComparison.CurrentCultureIgnoreCase)))
            {
                binTables.Add(failFlag.Replace("F_EVS", "Bin_EVS"));
                binTableEw.Add(failFlag.Replace("F_EVS", "Bin_EVS"), instRow.BinCutInstanceRow.BintableEnableWd);
            }
        }

        private void AddJudgeDeferRows(SubFlowSheet flow, List<string> failFlags)
        {
            flow.AddRow(new FlowRow { Opcode = OpCode.If, Parameter = string.Join("&&", failFlags.Distinct()) });
            flow.AddRow(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = "F_EVS_Defer" });
            flow.AddRow(new FlowRow { Opcode = OpCode.EndIf });
        }

        protected override string GetEvsNormalInstanceName(BinCutFinalInstanceRow row)
        {
            return $"EVS_{row.BinCutInstanceRow.Instance}_{row.GetVoltageType()}";
        }

        public string GetEvsRampPowerName(BinCutFinalInstanceRow row, EvsCondition condition)
        {
            return $"EVS_Static_Power_Ramp_Multi_{row.BinCutInstanceRow.Instance}_{condition.JobStage}_{condition.EvsNum}";
        }

        public string GetProfileInstanceName(BinCutFinalInstanceRow row, EvsCondition condition, string status)
        {
            return $"Profile_{status}_EVS_{row.BinCutInstanceRow.Instance}_{condition.JobStage}_{condition.EvsNum}";
        }

        public (InstanceRow instanceRow, string instancePinList, Function function) GenEvsRampInstance(BinCutFinalInstanceRow row, EvsCondition condition)
        {
            InstanceRow instanceRow = new InstanceRow();
            string instancePinList = "";
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameEvsRampPower, "evs", true);
            string selector = row.GetVoltageType();
            instanceRow.TestName = GetEvsRampPowerName(row, condition);
            instanceRow.TimeSets = row.PatSetName.Any() ? row.GetTimeSetVersion(row.PatternList) : "";
            instanceRow.VbtType = ".NET";
            instanceRow.DcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow.DCcategory) ? GetDcCategory(row) : row.BinCutInstanceRow.DCcategory.Split(' ').First();
            instanceRow.DcSelector = GetDcSelector(selector);
            instanceRow.AcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow.AcSpec) ? GetAcCategory(row.BinCutInstanceRow, instanceRow.TimeSets) : row.BinCutInstanceRow.AcSpec;
            instanceRow.AcSelector = "Typ";
            instanceRow.PinLevels = GenerateLevel(row.GetBlockByFlowName(), instanceRow.DcCategory, row.BinCutInstanceRow.Levels);
            instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
            instanceRow.SheetName = row.BinCutInstanceRow.SheetName;
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
                instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
            }
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;

            string evsParallelSettings = ApplyRampVoltageSettings(condition, function, ref instancePinList);
            function.SetParamValue("flagSerial", "FALSE");
            function.SetParamValue("dcSpec", string.IsNullOrEmpty(row.BinCutInstanceRow.DCcategory) ? GetDcCategory(row) : row.BinCutInstanceRow.DCcategory.Split(' ').First());
            function.SetParamValue("stressTimeSec", condition.Time);
            function.SetParamValue("stepNumber", "2");
            function.SetParamValue("risingDelayTimeSec", string.IsNullOrEmpty(condition.RisingDelayTimeSec) ? "0" : condition.RisingDelayTimeSec);
            function.SetParamValue("vTriggerControl", "FALSE");
            function.SetParamValue("vTriggerRange", "");
            function.SetParamValue("vTriggerIndexName", "");
            function.SetParamValue("vTriggerMaxStepName", "");
            function.SetParamValue("openLatchUpMeasure", "FALSE");
            function.SetParamValue("multiEVSIndex", condition.Pulses);
            function.SetParamValue("parallelPinVoltage", evsParallelSettings);
            int.TryParse(condition.Pulses, out int rampingCount);
            function.SetParamValue("multiFunction", rampingCount > 1 ? "TRUE" : "FALSE");
            function.SetParamValue("coolingTimeSec", condition.Cooling);
            function.SetParamValue("testTimeBreakdown", "TRUE");
            function.SetParamValue("totalPowerLimit", string.IsNullOrEmpty(condition.TotalPwr) ? "0" : condition.TotalPwr);
            if (row.BinCutInstanceRow.Evs.EvsType.Trim().ToLower() == "dynamic")
            {
                function.SetParamValue("dvsPlPatternSet", row.PatternList.Count == 1 ? row.PatternList[0] : row.PatSetName);
                function.SetParamValue("dvsCoolingPatIndex", row.BinCutInstanceRow.Evs.DvsCoolingPatIndex.Trim());
                function.SetParamValue("dvsCoolingPatSetLoopNumber", row.BinCutInstanceRow.Evs.DvsCoolingPatSetLoopNumber.Trim());
                bool isScan = row.BinCutInstanceRow.SubFlow.ContainsIgnoreCase("SCAN");
                string alarmFailFlag = condition.AlarmFlag != "" ? condition.AlarmFlag : "F_EVS_" + row.Domain + (isScan ? "_Sa_Alarm" : "_Bist_Alarm");
                function.SetParamValue("alarmFlag", alarmFailFlag.Trim());
            }

            if (condition.UserFunction.TryGetValue("Ramp", out string value) && !string.IsNullOrEmpty(value))
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(value, function);
            }

            instanceRow.Args = function.ArgList;
            return (instanceRow, instancePinList, function);
        }

        private string ApplyRampVoltageSettings(EvsCondition condition, Function function, ref string instancePinList)
        {
            string evsParallelSettings = "";
            if (!string.IsNullOrEmpty(condition.Voltage1))
            {
                List<string> pinSettings = condition.Voltage1.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                var pinSettingDic = new Dictionary<string, string>();
                foreach (string pinSetting in pinSettings)
                {
                    string[] spt = pinSetting.Split(':');
                    if (spt.Length == 2)
                    {
                        if (pinSettingDic.ContainsKey(spt[0]))
                        {
                            pinSettingDic[spt[0]] = pinSettingDic[spt[0]] + "," + spt[1];
                        }
                        else
                        {
                            pinSettingDic.Add(spt[0], spt[1]);
                        }
                    }
                }
                var pinSettings2 = new List<string>();
                var pinSettingDic2 = new Dictionary<string, string>();
                if (!string.IsNullOrEmpty(condition.Voltage2))
                {
                    pinSettings2 = condition.Voltage2.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    foreach (string pinSetting in pinSettings2)
                    {
                        string[] spt = pinSetting.Split(':');
                        if (spt.Length == 2)
                        {
                            if (pinSettingDic2.ContainsKey(spt[0]))
                            {
                                pinSettingDic2[spt[0]] = pinSettingDic2[spt[0]] + "," + spt[1];
                            }
                            else
                            {
                                pinSettingDic2.Add(spt[0], spt[1]);
                            }
                        }
                    }
                }
                var pinSettingList = pinSettingDic.Select(dic => dic.Key + ":" + dic.Value).ToList();
                if (pinSettingDic2.Any())
                {
                    pinSettingList.AddRange(pinSettingDic2.Select(dic => dic.Key + ":" + dic.Value).ToList());
                }

                evsParallelSettings = string.Join(";", pinSettingList);
                var pinList = pinSettings.Select(x => x.Split(':')[1]).ToList();
                List<string> pinList2 = pinSettings2.Any() ? pinSettings2.Select(x => x.Split(':')[1]).ToList() : null;
                if (pinList2 != null)
                {
                    function.SetParamValue("powerPin1", string.Join(",", pinList));
                    function.SetParamValue("powerPin2", string.Join(",", pinList2));
                    function.SetParamValue("paEVSEnable", "TRUE");
                    var fullPinList = new List<string>();
                    fullPinList.AddRange(pinList);
                    fullPinList.AddRange(pinList2);
                    function.SetParamValue("stressPowerPin", string.Join(",", fullPinList));
                    instancePinList = string.Join(",", fullPinList);
                }
                else
                {
                    function.SetParamValue("paEVSEnable", "FALSE");
                    function.SetParamValue("stressPowerPin", string.Join(",", pinList));
                    instancePinList = string.Join(",", pinList);
                }
            }
            return evsParallelSettings;
        }

        protected InstanceRow GenEvsCurrentProfileStartInstance(InstanceRow instanceRow, BinCutFinalInstanceRow row, string pinList, EvsCondition condition)
        {
            var profileStart = new InstanceRow
            {
                ColumnA = "Current profile start for " + instanceRow.ColumnA,
                TestName = GetProfileInstanceName(row, condition, "Start"),
                VbtType = ".NET"
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("ProfileAutoResolution", "evs", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenEvsCurrentProfileStartInstance(instanceRow, row.BinCutInstanceRow.SubFlow, pinList, "");
            }

            profileStart.VbtName = function.FullFunctionName;
            function.SetParamValue("pinName", pinList);
            function.SetParamValue("whatToCapture", "I");
            function.SetParamValue("capSignalName", "Capture_Signal");
            function.SetParamValue("plotTime", "10");
            function.SetParamValue("byFlow", "FALSE");
            profileStart.ArgList = function.Parameters;
            profileStart.Args = function.ArgList;
            return profileStart;
        }

        protected InstanceRow GenEvsCurrentProfilePlotInstance(InstanceRow instanceRow, BinCutFinalInstanceRow row, string pinList, EvsCondition condition)
        {
            var profilePlot = new InstanceRow
            {
                ColumnA = "Current profile plot for " + instanceRow.ColumnA,
                TestName = GetProfileInstanceName(row, condition, "Plot"),
                VbtType = ".NET"
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PlotProfile", "evs", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenEvsCurrentProfilePlotInstance(instanceRow, row.BinCutInstanceRow.SubFlow, pinList, "");
            }

            profilePlot.VbtName = function.FullFunctionName;
            function.SetParamValue("pinName", pinList);
            function.SetParamValue("capSignalName", "Capture_Signal");
            function.SetParamValue("exportWaveform", "TRUE");
            function.SetParamValue("plotWaveform", "FALSE");
            function.SetParamValue("calculateProfileInfo", "FALSE");
            function.SetParamValue("percentControl", "");
            profilePlot.ArgList = function.Parameters;
            profilePlot.Args = function.ArgList;
            return profilePlot;
        }
    }
}
