using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.BinCut.Base;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Atpg;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;

using LogLib.Static;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.DataStruct;
using TestPlanLib.EVS;

using Function = TestPlanLib.VbtLib.Function;

namespace Automation.GenerateIgxl.EVS
{
    public class EvsInstanceMainCs : EvsInstanceMain
    {
        public EvsInstanceMainCs(ScanConfig config, List<BinCutInstanceSheet> evsInstanceSheets) : base(config, evsInstanceSheets)
        {
        }

        internal override (InstanceRow instanceRow, string instancePinList, Function vbtFunctionBase) GenEVS_Static_Power_Ramp(BinCutFinalInstanceRow row)
        {
            Function function;
            var instanceRow = new InstanceRow();
            string selector = row.GetVoltageType();
            string instancePinList;
            instanceRow.TestName = row.GetEvsRampPowerName();
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

            if (EvsPowerPinIsFromTestPlan(row))
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameEvsRampPower, "evs");
                string pwrPinList = "";
                if (!string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsPwrPin1))
                {
                    string[] pwrPin1List = row.BinCutInstanceRow.Evs.EvsPwrPin1.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    string powerList = string.Join(",", pwrPin1List);
                    function.SetParamValue("powerPin1", powerList);
                    pwrPinList = powerList;
                    function.SetParamValue("paEVSEnable", "TRUE");
                }

                if (!string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsPwrPin2))
                {
                    string[] pwrPin2List = row.BinCutInstanceRow.Evs.EvsPwrPin2.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    string powerList = string.Join(",", pwrPin2List);
                    function.SetParamValue("powerPin2", powerList);
                    pwrPinList += "," + powerList;
                    function.SetParamValue("paEVSEnable", "TRUE");
                }
                if (!string.IsNullOrEmpty(pwrPinList))
                {
                    function.SetParamValue("stressPowerPin", pwrPinList);
                }

                instancePinList = pwrPinList;
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameEvsRampPower, "evs", true);
                if (!function.IsFound || function.Type == "VBT")
                {
                    return base.GenEVS_Static_Power_Ramp(row);
                }

                /* Compare dc/evs voltage to select power pin */
                string pwrPinList = BuildStressPowerPinList(row, instanceRow, selector);
                function.SetParamValue("stressPowerPin", pwrPinList);
                instancePinList = pwrPinList;
            }

            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            function.SetParamValue("flagSerial", string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsParallelSetting) ? "TRUE" : "FALSE");
            string evsParallelSettings = "";
            if (!string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsParallelSetting))
            {
                string[] pinSettings = row.BinCutInstanceRow.Evs.EvsParallelSetting.Split(';');
                var pinList = pinSettings.Select(x => x.Split(':')[1]).ToList();
                function.SetParamValue("stressPowerPin", string.Join(",", pinList));
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
            string evsDcSpec = row.BinCutInstanceRow.Evs.EvsCategory;
            if (string.IsNullOrEmpty(evsDcSpec))
            {
                evsDcSpec = string.IsNullOrEmpty(row.BinCutInstanceRow.DCcategory) ? GetDcCategory(row) : row.BinCutInstanceRow.DCcategory.Split(' ').First();
            }

            function.SetParamValue("dcSpec", evsDcSpec);
            function.SetParamValue("stressTimeSec", row.BinCutInstanceRow.Evs.EvsStressTime);
            function.SetParamValue("stepNumber", "2");
            function.SetParamValue("risingDelayTimeSec", "0");
            function.SetParamValue("vTriggerControl", "FALSE");
            function.SetParamValue("vTriggerRange", "");
            function.SetParamValue("vTriggerIndexName", "");
            function.SetParamValue("vTriggerMaxStepName", "");
            function.SetParamValue("openLatchUpMeasure", "FALSE");
            function.SetParamValue("multiEVSIndex", row.BinCutInstanceRow.Evs.EvsRampingCount);
            function.SetParamValue("parallelPinVoltage", evsParallelSettings);
            int.TryParse(row.BinCutInstanceRow.Evs.EvsRampingCount, out int rampingCount);
            function.SetParamValue("multiFunction", rampingCount > 1 ? "TRUE" : "FALSE");
            function.SetParamValue("coolingTimeSec", row.BinCutInstanceRow.Evs.EvsCoolingTime);
            function.SetParamValue("testTimeBreakdown", "TRUE");
            if (function.FunctionName.Equals(DcContiConst.CSharpFuncNameEvsRampPower))
            {
                function.SetParamValue("totalPowerLimit", string.IsNullOrEmpty(row.BinCutInstanceRow.Evs.EvsTotalPwrLimit) ? "0" : row.BinCutInstanceRow.Evs.EvsTotalPwrLimit);
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

        internal override InstanceRow GenEvsNormalInstance(BinCutFinalInstanceRow row)
        {
            bool isCall = !row.BinCutInstanceRow.PatternList.Any() && !row.BinCutInstanceRow.Evs.EvsConditions.Any(x => !string.IsNullOrEmpty(x.Voltage1));
            var instanceRow = new InstanceRow();
            Function function = new Function();
            string selector = row.GetVoltageType();
            instanceRow.TestName = isCall ? row.BinCutInstanceRow.Instance : GetEvsNormalInstanceName(row);
            instanceRow.TimeSets = row.GetTimeSetVersion(row.PatternList);
            instanceRow.VbtType = ".NET";
            instanceRow.DcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow.DCcategory) ? GetDcCategory(row) : row.BinCutInstanceRow.DCcategory.Split(' ').First();
            instanceRow.DcSelector = GetDcSelector(selector);
            instanceRow.AcCategory = string.IsNullOrEmpty(row.BinCutInstanceRow.AcSpec) ? GetAcCategory(row.BinCutInstanceRow, instanceRow.TimeSets) : row.BinCutInstanceRow.AcSpec;
            instanceRow.AcSelector = "Typ";
            instanceRow.PinLevels = GenerateLevel(row.GetBlockByFlowName(), instanceRow.DcCategory, row.BinCutInstanceRow.Levels);
            instanceRow.RowNum = row.BinCutInstanceRow.RowNum;
            instanceRow.SheetName = row.BinCutInstanceRow.SheetName;

            UserFunctionTableRow ufFuncSetting = null;
            if (!string.IsNullOrEmpty(row.BinCutInstanceRow.UserFunction) && TestPlanStatic.UserFunctionSheet != null)
            {
                ufFuncSetting = TestPlanStatic.UserFunctionSheet.Rows
                .FirstOrDefault(x => x.UserFunction.Equals(row.BinCutInstanceRow.UserFunction, StringComparison.OrdinalIgnoreCase));
            }

            if (row.BinCutInstanceRow != null && !string.IsNullOrEmpty(row.BinCutInstanceRow.SheetName))
            {
                instanceRow.ColumnA = row.BinCutInstanceRow.GetColumnA();
                instanceRow.ColumnA += ";DC Category:" + instanceRow.DcCategory;
            }
            if (row.BinCutInstanceRow.Instance.Equals("EVS_IrangeSetUp", StringComparison.CurrentCultureIgnoreCase))
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName("EVSIrangeSetup", "evs", true);
                if (!function.IsFound || function.Type == "VBT")
                {
                    return instanceRow;
                }
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameFunctionalT, "evs", true);
                if (!function.IsFound || function.Type == "VBT")
                {
                    return base.GenEvsNormalInstance(row);
                }

                function.SetParamValue("patterns", row.PatternList.Count == 1 ? row.PatternList[0] : row.PatSetName);
                function.SetParamValue("resultMode", row.IsBurstNonBinCutInstance() ? "1" : "0");
                function.SetParamValue("relayMode", "1");

                List<string> ufDigSrcPats = TestPlanStatic.UfDigSrcSheets
                    .SelectMany(x => x.Rows)
                    .Where(y => !string.IsNullOrEmpty(y.PatternName)).Select(z => z.PatternName).ToList();
                AtpgService.SetDigSrc(row, ufDigSrcPats, ufFuncSetting, LocalSpecs.HardIpInfos, "", row.PatternList, ref function);
            }
            if (ufFuncSetting != null)
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(row.BinCutInstanceRow.UserFunction, function);
            }
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        internal override InstanceRow GenEvsCurrentProfileStartInstance(InstanceRow instanceRow, string subFlowName, string pinList, string suffix)
        {
            var profileStart = new InstanceRow
            {
                ColumnA = "Current profile start for " + instanceRow.ColumnA,
                TestName = "Profile_Start_EVS_" + subFlowName + (subFlowName.ContainsIgnoreCase("SCAN") ? "_Scan" : "_Mbist")
            };
            if (!string.IsNullOrEmpty(suffix))
            {
                profileStart.TestName = profileStart.TestName + "_" + suffix;
            }

            profileStart.VbtType = ".NET";
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("ProfileAutoResolution", "evs", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenEvsCurrentProfileStartInstance(instanceRow, subFlowName, pinList, "");
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

        internal override InstanceRow GenPowerDown_EVS()
        {
            bool isUfpOnly = !TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlex);
            var powerDownEvs = new InstanceRow
            {
                ColumnA = "Power down instance for power reset flow",
                TestName = "PowerDown_EVS",
                VbtType = ".NET",
                DcCategory = "nWire_X_X_X",
                DcSelector = "Typ",
                AcCategory = isUfpOnly ? "" : "Common",
                AcSelector = isUfpOnly ? "" : "Typ",
                TimeSets = isUfpOnly ? "" : "TimeSet_nWire",
                PinLevels = "Levels_nWire"
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PowerDown", "evs", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenPowerDown_EVS();
            }

            powerDownEvs.VbtName = function.FullFunctionName;
            powerDownEvs.ArgList = function.Parameters;
            powerDownEvs.Args = function.ArgList;
            return powerDownEvs;
        }

        internal override InstanceRow GenPowerUp_EVS()
        {
            bool isUfpOnly = !TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlex);
            var powerUpEvs = new InstanceRow
            {
                ColumnA = "Power up instance for power reset flow",
                TestName = "PowerUp_EVS",
                VbtType = ".NET",
                DcCategory = "nWire_X_X_X",
                DcSelector = "Typ",
                AcCategory = isUfpOnly ? "" : "Common",
                AcSelector = isUfpOnly ? "" : "Typ",
                TimeSets = isUfpOnly ? "" : "TimeSet_nWire",
                PinLevels = "Levels_nWire"
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PowerUp", "evs", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenPowerUp_EVS();
            }

            function.SetParamValue("rampWaitTime", "0.001");
            function.SetParamValue("sequenceWaitTime", "0.001");
            function.SetParamValue("disconnectPins", "All_Digital_PowerUp");
            powerUpEvs.VbtName = function.FullFunctionName;
            powerUpEvs.ArgList = function.Parameters;
            powerUpEvs.Args = function.ArgList;
            return powerUpEvs;
        }

        protected override InstanceRow GenEvsCurrentProfilePlotInstance(InstanceRow instanceRow, string subFlowName, string pinList, string suffix)
        {
            var profilePlot = new InstanceRow
            {
                ColumnA = "Current profile plot for " + instanceRow.ColumnA,
                TestName = "Profile_Plot_EVS_" + subFlowName + (subFlowName.ContainsIgnoreCase("SCAN") ? "_Scan" : "_Mbist")
            };
            if (!string.IsNullOrEmpty(suffix))
            {
                profilePlot.TestName = profilePlot.TestName + "_" + suffix;
            }

            profilePlot.VbtType = ".NET";
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PlotProfile", "evs", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenEvsCurrentProfilePlotInstance(instanceRow, subFlowName, pinList, "");
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

        protected override InstanceRow GenEvsIvCurveInstance(InstanceRow instanceRow, Function vbtFunctionBase, EvsCondition condition)
        {
            InstanceRow copy = instanceRow.Copy();
            Function vbtFuncBaseIvCurve = vbtFunctionBase.Copy();
            copy.ColumnA = "IV curve for " + instanceRow.ColumnA;
            copy.TestName += "_IV";
            vbtFuncBaseIvCurve.SetParamValue("stressTimeSec", "0.001");
            vbtFuncBaseIvCurve.SetParamValue("stepNumber", "40");
            vbtFuncBaseIvCurve.SetParamValue("vTriggerControl", "FALSE");
            vbtFuncBaseIvCurve.SetParamValue("openLatchUpMeasure", "TRUE");
            vbtFuncBaseIvCurve.SetParamValue("multiFunction", "FALSE");
            vbtFuncBaseIvCurve.SetParamValue("multiEVSIndex", "");
            vbtFuncBaseIvCurve.SetParamValue("flagSerial", "TRUE");
            vbtFuncBaseIvCurve.SetParamValue("parallelPinVoltage", "");
            vbtFuncBaseIvCurve.SetParamValue("testTimeBreakdown", "FALSE");
            if (condition != null &&
                condition.UserFunction != null &&
                condition.UserFunction.TryGetValue("IV", out string value) &&
                !string.IsNullOrEmpty(value))
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(value, vbtFuncBaseIvCurve);
            }
            copy.Args = vbtFuncBaseIvCurve.ArgList;

            return copy;
        }

        internal override InstanceRow GenEvsVTrigInstance(InstanceRow instanceRow, Function vbtFunctionBase, EvsCondition condition)
        {
            InstanceRow copy = instanceRow.Copy();
            Function vbtFuncBaseVTrig = vbtFunctionBase.Copy();
            copy.ColumnA = "Vtrig for " + instanceRow.ColumnA;
            copy.TestName += "_Vtrig";
            vbtFuncBaseVTrig.SetParamValue("stressTimeSec", "0.2");
            vbtFuncBaseVTrig.SetParamValue("stepNumber", "2");
            vbtFuncBaseVTrig.SetParamValue("vTriggerControl", "TRUE");
            vbtFuncBaseVTrig.SetParamValue("vTriggerRange", "0.1");
            vbtFuncBaseVTrig.SetParamValue("vTriggerIndexName", "EVS_INDEX");
            vbtFuncBaseVTrig.SetParamValue("vTriggerMaxStepName", "EVS_Max_Step");
            vbtFuncBaseVTrig.SetParamValue("multiFunction", "FALSE");
            vbtFuncBaseVTrig.SetParamValue("multiEVSIndex", "");
            vbtFuncBaseVTrig.SetParamValue("flagSerial", "TRUE");
            vbtFuncBaseVTrig.SetParamValue("parallelPinVoltage", "");
            vbtFuncBaseVTrig.SetParamValue("testTimeBreakdown", "FALSE");
            if (condition != null &&
                condition.UserFunction != null &&
                condition.UserFunction.TryGetValue("Vtrig", out string value) &&
                !string.IsNullOrEmpty(value))
            {
                TestPlanStatic.UserFunctionSheet.ArgumentSetting(value, vbtFuncBaseVTrig);
            }
            copy.Args = vbtFuncBaseVTrig.ArgList;

            return copy;
        }

        protected virtual string GetEvsNormalInstanceName(BinCutFinalInstanceRow row)
        {
            return row.GetEvsParameter();
        }
    }
}
