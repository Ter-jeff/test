using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.GenerateIgxl.BinCut.Business;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.Scan.Harvest;
using Automation.GenerateIgxl.Scan.NonBinCut;
using Automation.Reader;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinNumber;
using TestPlanLib.DataStruct;
using TestPlanLib.EVS;
using TestPlanLib.Singleton;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.EVS
{
    public class EvsInstanceMain : ScanNonBinCutInstanceMain
    {
        protected readonly List<BinCutInstanceSheet> EvsInstanceSheets;

        private const string BinTableEvs = "Bin_Table_EVS";
        public const string FlowSa = "Flow_EVS_Sa";
        public const string FlowMbist = "Flow_EVS_Mbist";
        private List<string> _evsFailFlag = new List<string>();

        public EvsInstanceMain(ScanConfig config, List<BinCutInstanceSheet> evsInstanceSheets) : base(config)
        {
            EvsInstanceSheets = evsInstanceSheets;
        }

        internal List<InstanceRow> GenEvsInstances(IEnumerable<BinCutFinalInstanceRow> binCutFinalInstanceRows)
        {
            var instanceRows = new InstanceRows();
            var groups = binCutFinalInstanceRows.GroupBy(x => x.GetBlockByFlowName()).ToList();
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                foreach (BinCutFinalInstanceRow row in group)
                {
                    if (row.BinCutInstanceRow == null)
                    {
                        continue;
                    }

                    if ((!string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsCategory) || !string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsParallelSetting))
                        && row.BinCutInstanceRow.PatternList.Count > 0)
                    {
                        instanceRows.AddRange(GenEvsNormaInstanceAndRampInstance(row));
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsCategory) && string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsParallelSetting))
                        {
                            instanceRows.Add(GenEvsNormalInstance(row));
                        }
                        else
                        {
                            instanceRows.AddRange(GenEvsInstance(row));
                        }
                    }
                }
                if (group.Key.Equals("scan", StringComparison.CurrentCultureIgnoreCase))
                {
                    instanceRows.AddHeaderFooter(FlowSa);
                }
                else if (group.Key.Equals("mbist", StringComparison.CurrentCultureIgnoreCase))
                {
                    instanceRows.AddHeaderFooter(FlowMbist);
                }
            }
            instanceRows.AddRange(GenEvsPowerReset());
            return instanceRows;
        }

        private List<InstanceRow> GenEvsPowerReset()
        {
            var instances = new InstanceRows();
            if (TestPlanStatic.UfInstanceTable != null && !TestPlanStatic.UfInstanceTable.CheckExist("PowerUp_EVS"))
            {
                InstanceRow powerUpEvs = GenPowerUp_EVS();
                instances.Add(powerUpEvs);
            }
            if (TestPlanStatic.UfInstanceTable != null && !TestPlanStatic.UfInstanceTable.CheckExist("PowerDown_EVS"))
            {
                InstanceRow powerDownEvs = GenPowerDown_EVS();
                instances.Add(powerDownEvs);
            }
            return instances;
        }

        internal virtual InstanceRow GenPowerDown_EVS()
        {
            bool isUfpOnly = !TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlex);
            var powerDownEvs = new InstanceRow
            {
                ColumnA = "Power down instance for power reset flow",
                TestName = "PowerDown_EVS",
                VbtType = "VBT",
                DcCategory = "nWire_X_X_X",
                DcSelector = "Typ",
                AcCategory = isUfpOnly ? "" : "Common",
                AcSelector = isUfpOnly ? "" : "Typ",
                TimeSets = isUfpOnly ? "" : "TimeSet_nWire",
                PinLevels = "Levels_nWire"
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PowerDown_Parallel", "evs");
            function.SetParamValue("PowerPinList_DCVS", "All_Power");
            function.SetParamValue("DisconnectPinList", "All_Digital_PowerUp");
            function.SetParamValue("WaitConnectTime", "0.001");
            function.SetParamValue("DebugFlag", "-1");
            powerDownEvs.VbtName = function.FunctionName;
            powerDownEvs.ArgList = function.Parameters;
            powerDownEvs.Args = function.ArgList;
            return powerDownEvs;
        }

        internal virtual InstanceRow GenPowerUp_EVS()
        {
            bool isUfpOnly = !TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlex);
            var powerUpEvs = new InstanceRow
            {
                ColumnA = "Power up instance for power reset flow",
                TestName = "PowerUp_EVS",
                VbtType = "VBT",
                DcCategory = "nWire_X_X_X",
                DcSelector = "Typ",
                AcCategory = isUfpOnly ? "" : "Common",
                AcSelector = isUfpOnly ? "" : "Typ",
                TimeSets = isUfpOnly ? "" : "TimeSet_nWire",
                PinLevels = "Levels_nWire"
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PowerUp_Parallel", "evs");
            function.SetParamValue("PowerPinList_DCVS", "All_Power");
            function.SetParamValue("IO_H", "InitHi_Pins");
            function.SetParamValue("IO_L", "InitLo_Pins");
            function.SetParamValue("DisconnectPinList", "All_Digital_PowerUp");
            function.SetParamValue("WaitConnectTime", "0.001");
            function.SetParamValue("DebugFlag", "-1");
            powerUpEvs.VbtName = function.FunctionName;
            powerUpEvs.ArgList = function.Parameters;
            powerUpEvs.Args = function.ArgList;
            return powerUpEvs;
        }

        private List<InstanceRow> GenEvsNormaInstanceAndRampInstance(BinCutFinalInstanceRow row)
        {
            var evsInstances = new InstanceRows { GenEvsNormalInstance(row) };
            evsInstances.AddRange(GenEvsInstance(row));
            return evsInstances;
        }

        internal virtual InstanceRow GenEvsNormalInstance(BinCutFinalInstanceRow row)
        {
            var instanceRow = new InstanceRow();
            string selector = row.GetVoltageType();
            instanceRow.TestName = row.GetEvsParameter();
            instanceRow.TimeSets = row.GetTimeSetVersion(row.PatternList);
            instanceRow.VbtType = "VBT";
            instanceRow.DcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow.DCcategory) ? GetDcCategory(row) : row.BinCutInstanceRow.DCcategory.Split(' ').First();
            instanceRow.DcSelector = GetDcSelector(selector);
            instanceRow.AcCategory = GetAcCategory(row.BinCutInstanceRow, instanceRow.TimeSets);
            instanceRow.AcSelector = "Typ";
            instanceRow.PinLevels = GenerateLevel(row.GetBlockByFlowName(), instanceRow.DcCategory, row.BinCutInstanceRow.Levels);
            instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
            instanceRow.SheetName = row.BinCutInstanceRow.SheetName;
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
                instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
            }
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameFunctionalT, "evs");
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
            function.SetParamValue("Patterns", row.PatternList.Count == 1 ? row.PatternList[0] : row.PatSetName);
            function.SetParamValue("ResultMode", row.IsBurstNonBinCutInstance() ? "1" : "0");

            function.SetParamValue("RelayMode", "1");
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
                function.SetParamValue("DigSource", "Test_AutoSwitch:" + sendPinName.ToUpper());

            }
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction))
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, function);
            }

            instanceRow.VbtName = function.FunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        protected bool EvsPowerPinIsFromTestPlan(BinCutFinalInstanceRow row)
        {
            return row.BinCutInstanceRow != null &&
                   (!string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsPwrPin1) ||
                    !string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsPwrPin2));
        }

        protected virtual List<InstanceRow> GenEvsInstance(BinCutFinalInstanceRow row)
        {
            var instanceRows = new List<InstanceRow>();
            InstanceRow instanceRow;
            string instancePinList;
            Function vbtFunctionBase;
            (instanceRow, instancePinList, vbtFunctionBase) = GenEVS_Static_Power_Ramp(row);
            instanceRows.Add(instanceRow);
            instanceRows.Add(GenEvsVTrigInstance(instanceRow, vbtFunctionBase, null));
            instanceRows.Add(GenEvsIvCurveInstance(instanceRow, vbtFunctionBase, null));
            instanceRows.Add(GenEvsCurrentProfileStartInstance(instanceRow, row.BinCutInstanceRow.SubFlow, instancePinList, ""));
            instanceRows.Add(GenEvsCurrentProfilePlotInstance(instanceRow, row.BinCutInstanceRow.SubFlow, instancePinList, ""));

            return instanceRows;
        }

        internal virtual (InstanceRow instanceRow, string instancePinList, Function vbtFunctionBase) GenEVS_Static_Power_Ramp(BinCutFinalInstanceRow row)
        {
            Function function;
            var instanceRow = new InstanceRow();
            string selector = row.GetVoltageType();
            string instancePinList;
            instanceRow.TestName = row.GetEvsRampPowerName();
            instanceRow.TimeSets = row.PatSetName.Any() ? row.GetTimeSetVersion(row.PatternList) : "";
            instanceRow.VbtType = "VBT";
            instanceRow.DcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow.DCcategory) ? GetDcCategory(row) : row.BinCutInstanceRow.DCcategory.Split(' ').First();
            instanceRow.DcSelector = GetDcSelector(selector);
            instanceRow.AcCategory = GetAcCategory(row.BinCutInstanceRow, instanceRow.TimeSets);
            instanceRow.AcSelector = "Typ";
            instanceRow.PinLevels = GenerateLevel(row.GetBlockByFlowName(), instanceRow.DcCategory, row.BinCutInstanceRow.Levels);
            instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
            instanceRow.SheetName = row.BinCutInstanceRow.SheetName;
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
                instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
            }

            if (EvsPowerPinIsFromTestPlan(row))
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameEvsRampPowerPa, "evs");
                string pwrPinList = "";
                if (!string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsPwrPin1))
                {
                    string[] pwrPin1List = row.BinCutInstanceRow.Evs.EvsPwrPin1.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    string powerList = string.Join(",", pwrPin1List);
                    function.SetParamValue("PowerPin1", powerList);
                    pwrPinList = powerList;
                    function.SetParamValue("PA_Enable", "FALSE");
                }

                if (!string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsPwrPin2))
                {
                    string[] pwrPin2List = row.BinCutInstanceRow.Evs.EvsPwrPin2.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    string powerList = string.Join(",", pwrPin2List);
                    function.SetParamValue("PowerPin2", powerList);
                    pwrPinList += "," + powerList;
                    function.SetParamValue("PA_Enable", "TRUE");
                }
                if (!string.IsNullOrEmpty(pwrPinList))
                {
                    function.SetParamValue("Power_pin", pwrPinList);
                }

                instancePinList = pwrPinList;
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameEvsRampPower, "evs");

                /* Compare dc/evs voltage to select power pin */
                string pwrPinList = BuildStressPowerPinList(row, instanceRow, selector);
                function.SetParamValue("Power_pin", pwrPinList);
                instancePinList = pwrPinList;
            }

            instanceRow.VbtName = function.FunctionName;
            instanceRow.ArgList = function.Parameters;
            function.SetParamValue("Flag_Serial", string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsParallelSetting) ? "TRUE" : "FALSE");
            string evsParallelSettings = "";
            if (!string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsParallelSetting))
            {
                string[] pinSettings = row.BinCutInstanceRow.Evs.EvsParallelSetting.Split(';');

                var pinList = pinSettings.Select(x => x.Split(':')[1]).ToList();
                function.SetParamValue("Power_pin", string.Join(",", pinList));

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
                var pinSettingList = pinSettingDic.Select(dic => dic.Key + ":" + dic.Value).ToList();
                evsParallelSettings = string.Join(";", pinSettingList);
                instancePinList = string.Join(",", pinList);
            }
            function.SetParamValue("Parallel_Pin_Voltage", evsParallelSettings);
            function.SetParamValue("dc_spec", row.BinCutInstanceRow.Evs.EvsCategory);
            function.SetParamValue("S_WaitTime", row.BinCutInstanceRow.Evs.EvsStressTime);
            function.SetParamValue("timeset", row.BinCutInstanceRow.TimeSet);
            function.SetParamValue("Step_number", "2");
            function.SetParamValue("Rising_Delay_time", "0");
            function.SetParamValue("Looping_Contorl", "FALSE");
            function.SetParamValue("Looping_Range", "");
            function.SetParamValue("Looping_Index_Name", "");
            function.SetParamValue("Looping_Max_Steps_Name", "");
            function.SetParamValue("Open_LatchUp_measure", "FALSE");
            function.SetParamValue("Mulit_EVS_Index", row.BinCutInstanceRow.Evs.EvsRampingCount);
            int.TryParse(row.BinCutInstanceRow.Evs.EvsRampingCount, out int rampingCount);
            function.SetParamValue("Multi_Function", rampingCount > 1 ? "TRUE" : "FALSE");
            function.SetParamValue("Cooling_Time", row.BinCutInstanceRow.Evs.EvsCoolingTime);
            function.SetParamValue("Test_time_breakdown", "TRUE");
            if (function.FunctionName.Equals(DcContiConst.VbtFuncNameEvsRampPower))
            {
                function.SetParamValue("TotalPWRLimit", string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsTotalPwrLimit) ? "0" : row.BinCutInstanceRow.Evs.EvsTotalPwrLimit);
            }
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction))
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, function);
            }

            instanceRow.Args = function.ArgList;

            return (instanceRow, instancePinList, function);
        }

        private string BuildStressPowerPinList(BinCutFinalInstanceRow row, InstanceRow instanceRow, string selector)
        {
            List<TestSettingRow> evsDc = MultiTestSettingSheetsSingleton.Instance().TestSettingSheetsList[0].DataRows;
            string pwrPinList = "";
            foreach (TestSettingRow testSettingRow in evsDc)
            {
                if (string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsParallelSetting))
                {
                    if (testSettingRow.PowerPinName.Contains("_VOP") || testSettingRow.PowerPinName.Contains("_Valt"))
                    {
                        continue;
                    }

                    int evsVoltage = -1;
                    if (testSettingRow.DcCategoryValues.Exists(x => x.CategoryName.Equals(row.BinCutInstanceRow.Evs.EvsCategory, StringComparison.InvariantCultureIgnoreCase)))
                    {
                        int.TryParse(testSettingRow.DcCategoryValues.Find(x => x.CategoryName.Equals(row.BinCutInstanceRow.Evs.EvsCategory, StringComparison.InvariantCultureIgnoreCase)).Hv.OriginValue, out evsVoltage);
                    }

                    int dcVoltage = -1;
                    switch (selector)
                    {
                        case "HV":
                            {
                                if (testSettingRow.DcCategoryValues.Exists(x => x.CategoryName.Equals(instanceRow.DcCategory, StringComparison.InvariantCultureIgnoreCase)))
                                {
                                    int.TryParse(testSettingRow.DcCategoryValues.Find(x => x.CategoryName.Equals(instanceRow.DcCategory, StringComparison.InvariantCultureIgnoreCase)).Hv.OriginValue, out dcVoltage);
                                }

                                break;
                            }
                        case "LV":
                            {
                                if (testSettingRow.DcCategoryValues.Exists(x => x.CategoryName.Equals(instanceRow.DcCategory)))
                                {
                                    int.TryParse(testSettingRow.DcCategoryValues.Find(x => x.CategoryName.Equals(instanceRow.DcCategory)).Lv.OriginValue, out dcVoltage);
                                }

                                break;
                            }
                        case "NV":
                            {
                                if (testSettingRow.DcCategoryValues.Exists(x => x.CategoryName.Equals(instanceRow.DcCategory, StringComparison.InvariantCultureIgnoreCase)))
                                {
                                    int.TryParse(testSettingRow.DcCategoryValues.Find(x => x.CategoryName.Equals(instanceRow.DcCategory, StringComparison.InvariantCultureIgnoreCase)).Nv.OriginValue, out dcVoltage);
                                }

                                break;
                            }
                    }
                    if (evsVoltage == -1)
                    {
                        string alarmStr = $"Category {row.BinCutInstanceRow.Evs.EvsCategory} does not have HV value";
                        Response.Report(alarmStr, EnumMessageLevel.Error, 100);
                    }
                    if (dcVoltage == -1)
                    {
                        string alarmStr = $"Category {instanceRow.DcCategory} does not have {selector} value";
                        Response.Report(alarmStr, EnumMessageLevel.Error, 100);
                    }
                    if (evsVoltage > dcVoltage)
                    {
                        if (pwrPinList == "")
                        {
                            pwrPinList = testSettingRow.PowerPinName;
                        }
                        else
                        {
                            pwrPinList += "," + testSettingRow.PowerPinName;
                        }
                    }
                }
                if (pwrPinList == "")
                {
                    pwrPinList = testSettingRow.PowerPinName;
                }
                else
                {
                    pwrPinList += "," + testSettingRow.PowerPinName;
                }
            }
            return pwrPinList;
        }

        internal virtual InstanceRow GenEvsCurrentProfileStartInstance(InstanceRow instanceRow, string subFlowName, string pinList, string suffix)
        {
            var profileStart = new InstanceRow
            {
                ColumnA = "Current profile start for " + instanceRow.ColumnA,
                TestName = "Profile_Start_EVS_" + subFlowName + (subFlowName.ContainsIgnoreCase("SCAN") ? "_Scan" : "_Mbist"),
                VbtType = "VBT",
                VbtName = "Start_Profile_AutoResolution"
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Start_Profile_AutoResolution", "evs");
            function.SetParamValue("PinName", pinList);
            function.SetParamValue("WhatToCapture", "I");
            function.SetParamValue("CapSignalName", "Capture_Signal");
            function.SetParamValue("Plottime", "10");
            function.SetParamValue("ByFlow", "FALSE");
            profileStart.ArgList = function.Parameters;
            profileStart.Args = function.ArgList;

            return profileStart;
        }

        protected virtual InstanceRow GenEvsCurrentProfilePlotInstance(InstanceRow instanceRow, string subFlowName, string pinList, string suffix)
        {
            var profilePlot = new InstanceRow
            {
                ColumnA = "Current profile plot for " + instanceRow.ColumnA,
                TestName = "Profile_Plot_EVS_" + subFlowName + (subFlowName.ContainsIgnoreCase("SCAN") ? "_Scan" : "_Mbist"),
                VbtType = "VBT",
                VbtName = "Plot_Profile"
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Plot_Profile", "evs");
            function.SetParamValue("PinName", pinList);
            function.SetParamValue("CapSignalName", "Capture_Signal");
            function.SetParamValue("ExportWaveform", "TRUE");
            function.SetParamValue("PlotWaveform", "FALSE");
            function.SetParamValue("Calculate_ProfileInfo", "FALSE");
            profilePlot.ArgList = function.Parameters;
            profilePlot.Args = function.ArgList;

            return profilePlot;
        }

        protected virtual InstanceRow GenEvsIvCurveInstance(InstanceRow instanceRow, Function vbtFunctionBase, EvsCondition condition)
        {
            InstanceRow copy = instanceRow.Copy();
            Function function = vbtFunctionBase.Copy();
            copy.ColumnA = "IV curve for " + instanceRow.ColumnA;
            copy.TestName += "_IV";
            function.SetParamValue("S_WaitTime", "0.001");
            function.SetParamValue("Step_number", "40");
            function.SetParamValue("Open_LatchUp_measure", "TRUE");
            function.SetParamValue("Multi_Function", "FALSE");
            function.SetParamValue("Mulit_EVS_Index", "");
            function.SetParamValue("Flag_Serial", "TRUE");
            function.SetParamValue("Parallel_Pin_Voltage", "");
            function.SetParamValue("Test_time_breakdown", "FALSE");
            copy.Args = function.ArgList;

            return copy;
        }

        internal virtual InstanceRow GenEvsVTrigInstance(InstanceRow instanceRow, Function vbtFunctionBase, EvsCondition condition)
        {
            InstanceRow copy = instanceRow.Copy();
            Function function = vbtFunctionBase.Copy();
            copy.ColumnA = "Vtrig for " + instanceRow.ColumnA;
            copy.TestName += "_Vtrig";
            function.SetParamValue("S_WaitTime", "0.2");
            function.SetParamValue("Step_number", "2");
            function.SetParamValue("Looping_Contorl", "TRUE");
            function.SetParamValue("Looping_Range", "0.1");
            function.SetParamValue("Looping_Index_Name", "EVS_INDEX");
            function.SetParamValue("Looping_Max_Steps_Name", "EVS_Max_Step");
            function.SetParamValue("Multi_Function", "FALSE");
            function.SetParamValue("Mulit_EVS_Index", "");
            function.SetParamValue("Flag_Serial", "TRUE");
            function.SetParamValue("Parallel_Pin_Voltage", "");
            function.SetParamValue("Test_time_breakdown", "FALSE");
            copy.Args = function.ArgList;

            return copy;
        }

        protected void GenEvsInstanceSheet(List<InstanceRow> instanceRows)
        {
            if (instanceRows.Count > 0)
            {
                var instanceSheet = new InstanceSheet("TestInst_EVS");
                instanceSheet.AddRows(instanceRows);
                instanceSheet.RemoveDuplicateInstance(false);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirEvs, instanceSheet);
            }
        }

        internal override string GenerateLevel(string evsType, string dcCategory, string userDefinedLevel)
        {
            if (!string.IsNullOrEmpty(userDefinedLevel))
            {
                return userDefinedLevel;
            }

            string level = "";
            switch (evsType.ToLower())
            {
                case "scan":
                    {
                        level = "Levels_EVS_Scan";
                        break;
                    }
                case "mbist":
                    {
                        level = "Levels_EVS_Mbist";
                        if (!string.IsNullOrEmpty(dcCategory) && dcCategory.ContainsIgnoreCase("logic"))
                        {
                            level = "Levels_EVS_Scan";
                        }

                        break;
                    }
            }
            return level;
        }

        internal List<SubFlowSheet> GenEvsFlows(List<BinCutFinalInstanceRow> tpInstRows)
        {
            var flowSheets = new List<SubFlowSheet>();
            IEnumerable<IGrouping<string, BinCutFinalInstanceRow>> groupBySubflowName = tpInstRows.GroupBy(x => x.BinCutInstanceRow.SubFlow);
            var flowSheet = new SubFlowSheet("Flow_EVS");
            flowSheet.AddRow(new FlowRow { Opcode = "assign-integer", Parameter = "EVS_Max_Step 11" });
            flowSheet.AddRow(new FlowRow { Opcode = "assign-integer", Enable = "Vtrig && CurrentProfile", Parameter = "EVS_Max_Step 11" });
            flowSheet.AddRow(new FlowRow { Opcode = "assign-integer", Enable = "Vtrig && !CurrentProfile", Parameter = "EVS_Max_Step 11" });
            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groupBySubflowName)
            {
                if (group.Key.ContainsIgnoreCase("SCAN"))
                {
                    SubFlowSheet scanFlow = GenEvsScanFlow(group.ToList());
                    if (scanFlow.Rows.Any())
                    {
                        scanFlow.AddRow(new FlowRow { Opcode = "Next" });

                        scanFlow.AddEndRows();
                        flowSheets.Add(scanFlow);
                        flowSheet.AddRow(new FlowRow { Opcode = "call", Parameter = scanFlow.Name });
                    }
                }
                else if (group.Key.ContainsIgnoreCase("BIST") || group.Key.ContainsIgnoreCase("BIRA"))
                {
                    SubFlowSheet bistFlow = GenEvsMbistFlow(group.ToList());
                    if (bistFlow.Rows.Any())
                    {
                        bistFlow.AddRow(new FlowRow { Opcode = "Next" });

                        bistFlow.AddEndRows();
                        flowSheets.Add(bistFlow);
                        flowSheet.AddRow(new FlowRow { Opcode = "call", Parameter = bistFlow.Name });
                    }
                }
            }

            if (flowSheet.Rows.Any())
            {
                flowSheet.AddReturnRow();
                flowSheets.Add(flowSheet);
            }

            flowSheets.Add(GenPowerResetFlow());

            return flowSheets;
        }

        public SubFlowSheet GenPowerResetFlow(List<BinCutInstanceSheet> powerRestInstRows)
        {
            var powerResetFlow = new SubFlowSheet("Flow_EVS_PowerReset");
            if (powerRestInstRows == null || !powerRestInstRows.Any())
            {
                powerResetFlow.AddRow(new FlowRow { Opcode = "Test", Parameter = "PowerDown_EVS" });
                AddFreeRunClkInstance(ref powerResetFlow, true);
                powerResetFlow.AddRow(new FlowRow { Opcode = "Test", Parameter = "PowerUp_EVS" });
                powerResetFlow.AddRow(new FlowRow { Opcode = "Test", Parameter = "Relay_ON_Default" });
                AddFreeRunClkInstance(ref powerResetFlow);
            }
            else
            {
                foreach (BinCutInstanceSheet rows in powerRestInstRows)
                {
                    foreach (BinCutInstanceRow row in rows.Rows)
                    {
                        if (!string.IsNullOrEmpty(row.Instance))
                        {
                            string[] parts = row.Instance.Split('=');

                            if (parts.Length >= 2 &&
                                !string.IsNullOrEmpty(parts[0]) &&
                                !string.IsNullOrEmpty(parts[1]))
                            {
                                string prefix = row.Instance.Contains("Disable") ? "FreeRunClk_Disable_" : "FreeRunClk_Enable_";
                                foreach (EnumEquipment equipment in TestPlanStatic.Equipments)
                                {
                                    ProtocolAwarePin pin = NwireSingleton.Instance().SettingInfo.NwirePins
                                        .Find(x => !string.IsNullOrEmpty(x.OutClk) &&
                                        row.Instance.Contains(x.OutClk, StringComparison.OrdinalIgnoreCase));

                                    if (pin == null)
                                    {
                                        powerResetFlow.AddRow(new FlowRow { ColumnA = row.Instance });
                                        continue;
                                    }

                                    string disableInstName = prefix + pin.CreatePinNameWithDiff() + "_" + (equipment.Equals(EnumEquipment.UltraFlex) ? "UF" : "UFP");
                                    string testerType = equipment == EnumEquipment.UltraFlex ? "UF" : "UFP";
                                    string job = "";
                                    if (TestPlanStatic.JobInfoSheet != null)
                                    {
                                        IEnumerable<string> jobs = TestPlanStatic.JobInfoSheet.Rows.Where(x => x.TesterType.Equals(testerType)).Select(x => x.JobName).ToList();
                                        if (jobs.Any())
                                        {
                                            job = string.Join(",", jobs);
                                        }
                                    }
                                    powerResetFlow.AddRow(new FlowRow { Opcode = !string.IsNullOrEmpty(row.Opcode) ? row.Opcode : "Test", Parameter = disableInstName, Job = job, Enable = row.EnableFlow });
                                }
                            }
                            else
                            {
                                powerResetFlow.AddRow(new FlowRow { Opcode = !string.IsNullOrEmpty(row.Opcode) ? row.Opcode : "Test", Parameter = row.Instance, Enable = row.EnableFlow });
                            }
                        }
                    }
                }
            }

            powerResetFlow.AddReturnRow();
            return powerResetFlow;
        }

        public SubFlowSheet GenPowerResetFlow()
        {
            var powerResetFlow = new SubFlowSheet("Flow_EVS_PowerReset");
            powerResetFlow.AddRow(new FlowRow { Opcode = "Test", Parameter = "PowerDown_EVS" });
            AddFreeRunClkInstance(ref powerResetFlow, true);
            powerResetFlow.AddRow(new FlowRow { Opcode = "Test", Parameter = "PowerUp_EVS" });
            powerResetFlow.AddRow(new FlowRow { Opcode = "Test", Parameter = "Relay_ON_Default" });
            AddFreeRunClkInstance(ref powerResetFlow);
            powerResetFlow.AddReturnRow();
            return powerResetFlow;
        }

        private void AddFreeRunClkInstance(ref SubFlowSheet powerResetFlow, bool isDisable = false)
        {
            string prefix = isDisable ? "FreeRunClk_Disable_" : "FreeRunClk_Enable_";
            foreach (EnumEquipment equipment in TestPlanStatic.Equipments)
            {
                foreach (ProtocolAwarePin pin in NwireSingleton.Instance().SettingInfo.NwirePins)
                {
                    string disableInstName = prefix + pin.CreatePinNameWithDiff() + "_" + (equipment.Equals(EnumEquipment.UltraFlex) ? "UF" : "UFP");
                    string testerType = equipment == EnumEquipment.UltraFlex ? "UF" : "UFP";
                    string job = "";
                    if (TestPlanStatic.JobInfoSheet != null)
                    {
                        IEnumerable<string> jobs = TestPlanStatic.JobInfoSheet.Rows.Where(x => x.TesterType.Equals(testerType)).Select(x => x.JobName).ToList();
                        if (jobs.Any())
                        {
                            job = string.Join(",", jobs);
                        }
                    }
                    powerResetFlow.AddRow(new FlowRow { Opcode = "Test", Parameter = disableInstName, Job = job });
                }
            }
        }

        private SubFlowSheet GenEvsScanFlow(List<BinCutFinalInstanceRow> saInstRows) => GenerateEvsFlow(saInstRows, "Sa_Alarm");

        private SubFlowSheet GenEvsMbistFlow(List<BinCutFinalInstanceRow> bistInstRows) => GenerateEvsFlow(bistInstRows, "Bist_Alarm");

        private SubFlowSheet GenerateEvsFlow(List<BinCutFinalInstanceRow> instRows, string alarmSuffix)
        {
            if (instRows == null || !instRows.Any())
            {
                return null;
            }

            var evsFlow = new SubFlowSheet(instRows.First().BinCutInstanceRow.SubFlow);
            var binTableList = new List<string>();

            evsFlow.AddStartRows();
            evsFlow.AddRow(new FlowRow { Opcode = "For", Parameter = "EVS_INDEX=0; EVS_INDEX< EVS_Max_Step; EVS_INDEX++" });

            IEnumerable<IGrouping<string, BinCutFinalInstanceRow>> groups = instRows.GroupBy(x => x.Domain);

            foreach (IGrouping<string, BinCutFinalInstanceRow> group in groups)
            {
                var groupRows = group.ToList();
                for (int i = 0; i < groupRows.Count; i++)
                {
                    BinCutFinalInstanceRow instRow = groupRows[i];
                    BinCutInstanceRow binCut = instRow.BinCutInstanceRow;
                    if (binCut == null)
                    {
                        continue;
                    }

                    string failFlag = !string.IsNullOrEmpty(binCut.FailFlag) ? binCut.FailFlag : $"F_EVS_{group.Key}{alarmSuffix}";

                    if (!binCut.BinOutStage.Equals("X", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (!_evsFailFlag.Any(x => x.Equals(failFlag, StringComparison.OrdinalIgnoreCase)))
                        {
                            _evsFailFlag.Add(failFlag);
                        }
                        binTableList.Add(failFlag.Replace("F_EVS", "Bin_EVS"));
                    }

                    if (!string.IsNullOrEmpty(instRow.BinCutInstanceRow.SiteVar))
                    {
                        string siteVar = instRow.BinCutInstanceRow.SiteVar;
                        List<FlowRow> ifFlowRows;
                        if ((!string.IsNullOrEmpty(instRow.BinCutInstanceRow.Evs.EvsCategory) || !string.IsNullOrEmpty(instRow.BinCutInstanceRow.Evs.EvsParallelSetting)) && instRow.BinCutInstanceRow.PatternList.Count > 0)
                        {
                            ifFlowRows = GetIfFlowRows(siteVar, GetEvsNormalAndRampTestRows(instRow, true, failFlag), null);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(instRow.BinCutInstanceRow.Evs.EvsCategory) ||
                                !string.IsNullOrEmpty(instRow.BinCutInstanceRow.Evs.EvsParallelSetting))
                            {
                                ifFlowRows = GetIfFlowRows(siteVar, GetEvsTestRow(instRow, true, failFlag), null);
                            }
                            else
                            {
                                ifFlowRows = GetIfFlowRows(siteVar, GetEvsTestRow(instRow, false, failFlag), null);
                            }
                        }
                        evsFlow.AddRows(ifFlowRows);
                    }
                    else
                    {
                        if ((!string.IsNullOrEmpty(instRow.BinCutInstanceRow.Evs.EvsCategory) || !string.IsNullOrEmpty(instRow.BinCutInstanceRow.Evs.EvsParallelSetting)) && instRow.BinCutInstanceRow.PatternList.Count > 0)
                        {
                            evsFlow.AddRows(GetEvsNormalAndRampTestRows(instRow, false, failFlag));
                        }
                        else
                        {
                            evsFlow.AddRows(!string.IsNullOrEmpty(instRow.BinCutInstanceRow.Evs.EvsCategory) || !string.IsNullOrEmpty(instRow.BinCutInstanceRow.Evs.EvsParallelSetting) ? GetEvsTestRow(instRow, true, failFlag) : GetEvsTestRow(instRow, false));
                        }
                    }

                    if (i == group.Count() - 1)
                    {
                        foreach (string parameter in binTableList)
                        {
                            evsFlow.AddRow(new FlowRow { Opcode = OpCode.BinTable, Parameter = parameter });
                        }
                    }
                }
            }
            return evsFlow;
        }

        private List<FlowRow> GetIfFlowRows(string siteVar, List<FlowRow> testFlowRows, List<FlowRow> limitRows)
        {
            var ifFlowRows = new List<FlowRow>();
            var flowRow = new FlowRow();
            if (limitRows != null)
            {
                ifFlowRows.AddRange(limitRows);
            }

            foreach (FlowRow testFlowRow in testFlowRows)
            {
                if (!string.IsNullOrEmpty(siteVar))
                {
                    if (!string.IsNullOrEmpty(testFlowRow.DeviceName) ||
                        siteVar.Contains("&&") || siteVar.Contains("||"))
                    {
                        if (!testFlowRow.Enable.Equals("Vtrig"))
                        {
                            ifFlowRows.Clear();
                        }

                        ifFlowRows.Add(FlowRow.GenIfCondition(siteVar, testFlowRow.Job));
                        if (testFlowRows.Count != 1)
                        {
                            ifFlowRows.AddRange(testFlowRows);
                        }
                        else
                        {
                            ifFlowRows.Add(testFlowRow);
                        }

                        if (limitRows != null)
                        {
                            ifFlowRows.AddRange(limitRows);
                        }

                        ifFlowRows.Add(FlowRow.GenEndIf(testFlowRow.Job));
                        return ifFlowRows;
                    }

                    if (string.IsNullOrEmpty(testFlowRow.DeviceName))
                    {
                        testFlowRow.DeviceName = siteVar.TrimStart('!');
                        if (!siteVar.Trim().EndsWith("False", StringComparison.CurrentCultureIgnoreCase))
                        {
                            testFlowRow.DeviceCondition = siteVar.StartsWith("!") ? "Flag-false" : "Flag-true";
                        }
                        else
                        {
                            testFlowRow.DeviceCondition = siteVar.StartsWith("!") ? "Flag-true" : "Flag-false";
                        }
                    }
                }

                if (!string.IsNullOrEmpty(testFlowRow.DeviceName))
                {
                    var ifFlowRow = new FlowRow { Job = flowRow.Job, Opcode = "If", Parameter = siteVar };
                    ifFlowRows.Add(ifFlowRow);

                    testFlowRow.DeviceName = "";
                    testFlowRow.DeviceCondition = "";
                    ifFlowRows.Add(testFlowRow);

                    var endIfFlowRow = new FlowRow { Job = flowRow.Job, Opcode = "EndIf" };
                    ifFlowRows.Add(endIfFlowRow);
                }
            }
            return ifFlowRows;
        }

        private List<FlowRow> GetEvsNormalAndRampTestRows(BinCutFinalInstanceRow row, bool isSiteVar, string failFlag = null)
        {
            var flowRows = new List<FlowRow>();
            flowRows.AddRange(GetEvsTestRow(row, false, isSiteVar ? failFlag : null));
            flowRows.AddRange(GetEvsTestRow(row, true, failFlag));
            return flowRows;
        }

        public List<FlowRow> GenPowerResetRows(FlowRow flowRow)
        {
            var result = new List<FlowRow>();
            var ifFlowRow = new FlowRow { Opcode = OpCode.If, Parameter = "F_EVS_PwrDown" };
            result.Add(ifFlowRow);
            var enableFlowWordRow = new FlowRow { Job = flowRow.Job, Opcode = OpCode.EnableFlowWd, Parameter = "EVS_PwrDown" };
            result.Add(enableFlowWordRow);
            var endIfFlowRow = new FlowRow { Opcode = OpCode.EndIf };
            result.Add(endIfFlowRow);
            var callFlowRow = new FlowRow { Enable = "EVS_PwrDown", Job = flowRow.Job, Opcode = OpCode.Call, Parameter = "Flow_EVS_PowerReset" };
            result.Add(callFlowRow);
            var disableFlowWordRow = new FlowRow { Job = flowRow.Job, Opcode = OpCode.DisableFlowWd, Parameter = "EVS_PwrDown" };
            result.Add(disableFlowWordRow);

            return result;
        }

        private List<FlowRow> GetEvsTestRow(BinCutFinalInstanceRow row, bool isRamp, string failFlag = null)
        {
            var flowRow = new FlowRow();
            var flowRows = new List<FlowRow>();
            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                flowRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
            }

            flowRow.Opcode = row.NopByEnableWord ? OpCode.Nop : OpCode.Test;
            flowRow.Parameter = isRamp ? row.GetEvsRampPowerName() : row.GetEvsParameter();

            flowRow.Job = row.GetJob();
            flowRow.Enable = row.GetEnable();
            if (isRamp)
            {
                flowRow.Enable = string.IsNullOrEmpty(flowRow.Enable) ? "A_EVS" : flowRow.Enable + "&& A_EVS";
            }

            flowRow.FailAction = string.IsNullOrEmpty(failFlag) ? "" : failFlag;
            if (row.BinCutInstanceRow != null && row.BinCutInstanceRow.BinOutStage.Equals("X", StringComparison.CurrentCultureIgnoreCase) && !string.IsNullOrEmpty(flowRow.FailAction))
            {
                flowRow.FailAction += ", F_EVS_Tight";
            }

            flowRows.Add(flowRow);
            if (isRamp && !row.IsEvs2())
            {
                flowRows.Add(GenEvsVTrigTestRow(flowRow));
                flowRows.Add(GenEvsIvCurveTestRow(flowRow));
                flowRows.AddRange(GenPowerResetRows(flowRow));
            }
            return flowRows;
        }

        internal FlowRow GenEvsIvCurveTestRow(FlowRow flowRow)
        {
            FlowRow copy = flowRow.Copy();
            copy.ColumnA = "IV curve for " + flowRow.ColumnA;
            copy.Parameter = flowRow.Parameter + "_IV";
            copy.FailAction = copy.FailAction.Replace(", F_EVS_Tight", "").Trim();
            copy.Enable = "IVCurve";

            return copy;
        }

        internal FlowRow GenEvsVTrigTestRow(FlowRow flowRow)
        {
            FlowRow copy = flowRow.Copy();
            copy.ColumnA = "Vtrig for " + flowRow.ColumnA;
            copy.Parameter = flowRow.Parameter + "_Vtrig";
            copy.Enable = "Vtrig";

            return copy;
        }

        private BinTableSheet GenEvsBinTable()
        {
            var binTableSheet = new BinTableSheet(BinTableEvs);
            foreach (string flag in _evsFailFlag)
            {
                string bName = flag.Replace("F_EVS", "Bin_EVS");
                var bin = new BinTableRow { Name = bName, ItemList = flag, Op = "OR" };
                bin.Items.Add("T");
                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("EVS", "", "", bin);
                bin.Sort = binNumInfo.SoftBin.ToString("G15");
                bin.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                bin.Result = binNumInfo.BinNumInfo.Status;
                binTableSheet.AddRow(bin);
            }
            return binTableSheet;
        }

        public override void WorkFlow()
        {
            _evsFailFlag = new List<string>();
            BinTableSheet binTable;
            var patSets = new List<PatSet>();
            var flowSheets = new List<SubFlowSheet>();
            var instanceRows = new List<InstanceRow>();

            BinCutInstanceNamingSheet binCutInstanceNamingSheet = BinCutInstanceNamingSheet();
            var instSheetPreProcess = new InstSheetPreProcess(Config);
            var evsInstanceSheet = EvsInstanceSheets.FindAll(x => !x.SheetName.Equals("Instance_EVS_PowerReset")).ToList();
            var evsInstancePowerResetSheet = EvsInstanceSheets.FindAll(x => x.SheetName.Equals("Instance_EVS_PowerReset")).ToList();
            BinCutFinalInstanceRows tpInstRows = instSheetPreProcess.InitialInstance(evsInstanceSheet, binCutInstanceNamingSheet);
            tpInstRows = tpInstRows.RePatSetNameDuplicateRows();
            List<InstanceRow> instances;
            bool isMultiEvsFormat = evsInstanceSheet.First().Rows.Any(x => x.Evs.EvsConditions.Any());
            AcSpecSheet acSpecSheet = TestProgram.IgxlWorkBk.GetAcSpecsSheet();
            if (acSpecSheet != null)
            {
                new BinCutAcSpecsWriter().GenAcSpecs(tpInstRows, acSpecSheet, true);
            }
            //New format(Combine multiple job and evs1 & evs2)
            if (isMultiEvsFormat)
            {
                var evsMultiRampInstanceMain = new EvsMultiRampInstanceMain(Config, EvsInstanceSheets);
                instances = evsMultiRampInstanceMain.GenRampEvsInstances(tpInstRows);
                EvsInstanceNameCheck(ref tpInstRows, instances);

                // Gen EVS Instance
                instanceRows.AddRange(instances);
                // Gen EVS Flow
                flowSheets.AddRange(evsMultiRampInstanceMain.GenEvsFlow(tpInstRows));
                // Gen EVS Power Reset Flow
                flowSheets.Add(evsMultiRampInstanceMain.GenPowerResetFlow(evsInstancePowerResetSheet));
                // Gen BinTable
                binTable = evsMultiRampInstanceMain.GenEvsBinTable();
                // Gen PatSet
                patSets.AddRange(GenPatSets(tpInstRows));
            }
            //Original format
            else
            {
                if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
                {
                    instances = GenEvsInstances(tpInstRows);
                }
                else
                {
                    var evsInstanceMainCs = new EvsInstanceMainCs(Config, EvsInstanceSheets);
                    instances = evsInstanceMainCs.GenEvsInstances(tpInstRows);
                }
                EvsInstanceNameCheck(ref tpInstRows, instances);

                //Gen EVS Instance
                instanceRows.AddRange(instances);
                // Gen EVS Flow
                flowSheets.AddRange(GenEvsFlows(tpInstRows));
                // Gen BinTable
                binTable = GenEvsBinTable();
                // Gen PatSet
                patSets.AddRange(GenPatSets(tpInstRows));
            }
            #region add into igxl
            TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirEvs, binTable);
            SetPatSetSheet(patSets);
            foreach (SubFlowSheet flow in flowSheets)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirEvs, flow);
            }

            GenEvsInstanceSheet(instanceRows);
            #endregion
        }

        private void EvsInstanceNameCheck(ref BinCutFinalInstanceRows oldRows, List<InstanceRow> instRows)
        {
            for (int i = 0; i < instRows.Count; i++)
            {
                InstanceRow row1 = instRows[i];
                for (int j = i + 1; j < instRows.Count; j++)
                {
                    InstanceRow row2 = instRows[j];
                    if (row1.TestName.Equals(row2.TestName, StringComparison.CurrentCultureIgnoreCase) && row1.GetDifferences(row2, false).Any())
                    {
                        for (int m = 0; m < oldRows.Count; m++)
                        {
                            if (row2.RowNum.Equals(oldRows[m].BinCutInstanceRow.RowNum) &&
                                row2.SheetName.Equals(oldRows[m].BinCutInstanceRow.SheetName))
                            {
                                oldRows[m].IsDuplicateName = true;
                            }
                        }
                    }
                }
            }
        }

        internal override void SetPatSetSheet(List<PatSet> patSets)
        {
            if (patSets.Count > 0)
            {
                var patSetSheet = new PatSetSheet("PatSets_EVS");
                patSetSheet.AddRows(patSets);
                TestProgram.IgxlWorkBk.AddPatSetSheet(FolderStructure.DirEvs, patSetSheet);
            }
        }
    }
}
