using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenAc;
using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using TestPlanLib.BinNumber;
using TestPlanLib.DataStruct;
using TestPlanLib.Singleton;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public class ContiMain
    {
        protected DcTestContiSheet DcTestContiSheet;
        private readonly PatSetSheet _patSetAll;
        private readonly IoInfoSheet _ioInfoSheet;
        private readonly bool _needSetPowerAlarm;
        private readonly bool _needPowerUp;

        public ContiMain(DcTestContiSheet dcTestContiSheet, PatSetSheet patSetSheet, IoInfoSheet ioInfoSheet)
        {
            DcTestContiSheet = dcTestContiSheet;
            _patSetAll = patSetSheet;
            _ioInfoSheet = ioInfoSheet;
            _needSetPowerAlarm = !DcTestContiSheet.DcTestContiRows.Any(x => x.InstanceName.Equals("Disable_Set_Power_Alarm", StringComparison.OrdinalIgnoreCase));
            _needPowerUp = !DcTestContiSheet.DcTestContiRows.Any(x => x.InstanceName.Equals("Disable_Power_Up", StringComparison.OrdinalIgnoreCase));
        }

        public virtual ContiResult WorkFlow(MultiLevelSheets multiLevel, ref List<BinTableRow> binTableRows)
        {
            List<ContiBase> contiTestList = CreateMultiContiTest(DcTestContiSheet.DcTestContiRows);

            SubFlowSheet contiFlow = GenerateContiFlow(contiTestList);
            InstanceSheet contiInstanceSheet = GenerateContiInstanceSheet(contiTestList);
            List<InstanceRow> commonInstanceRows = GenerateCommonInstanceRows(contiTestList);
            GenerateBinTableRows(ref binTableRows, contiTestList);

            ContiResult result = new ContiResult { ContiFlowSheet = contiFlow, ConInstanceSheet = contiInstanceSheet };

            result.CommonInstance.Rows.AddRange(commonInstanceRows);

            CheckConti(contiTestList);

            var crwtTable = new CrwtTable(
                contiTestList,
                GetTesterTypeOrDefault("CP1")
            );
            crwtTable.ExportToTxt(FolderStructure.DirCommonSheets);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, "TuneCRWTResult");
            return result;

        }

        public static string GetTesterTypeOrDefault(
        string jobName,
        string defaultType = "UF")
        {

            if (TestPlanStatic.JobInfoSheet == null)
            {
                return defaultType;
            }

            string testerType = TestPlanStatic.JobInfoSheet.Rows
                       .Find(x => x.JobName.Equals(jobName, StringComparison.OrdinalIgnoreCase))?
                       .TesterType?
                       .ToUpperInvariant();

            return string.IsNullOrWhiteSpace(testerType) ? defaultType : testerType;
        }

        protected virtual void CheckConti(List<ContiBase> contiTestList)
        {
            foreach (ContiBase contiRow in contiTestList)
            {
                List<string> pinList = contiRow.DcTestContiRow.PinGroup.Split(';', ',').ToList();

                foreach (string pin in pinList)
                {
                    if (string.IsNullOrEmpty(pin))
                    {
                        continue;
                    }

                    if (TestProgram.IgxlWorkBk.PinMapPair.Value.GetGroup(pin) == null &&
                        TestProgram.IgxlWorkBk.PinMapPair.Value.GetPin(pin) == null)
                    {
                        ErrorReportManager.AddError(BasicErrorType.E_MissingPin_16, NeededSheets.ContiDcTest, 1, 1, $"Missing conti pin/pin group: {pin} in pin map!", new string[] { pin });
                    }
                }
            }
        }

        protected virtual List<ContiBase> CreateMultiContiTest(List<DcTestContiRow> dcTestContiData)
        {
            var contiTestList = new List<ContiBase>();
            foreach (DcTestContiRow dcTestContiRow in dcTestContiData)
            {
                switch (dcTestContiRow.TestType)
                {
                    case ContiType.OpenShort:
                        var openShortRow = new ContiOpenShort(dcTestContiRow);
                        contiTestList.Add(openShortRow);
                        break;
                    case ContiType.ContiAnalog:
                        ContiAnalog analogRow = !string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) ? new ContiAnalogCs(dcTestContiRow) : new ContiAnalog(dcTestContiRow);
                        contiTestList.Add(analogRow);
                        break;
                    case ContiType.ContiPpmu:
                        ContiPpmu contiPpmu = new ContiPpmu(dcTestContiRow);
                        contiTestList.Add(contiPpmu);
                        break;
                    case ContiType.WalkingZ:
                        var walkingZRow = new ContiWalkingZ(dcTestContiRow, _patSetAll);
                        contiTestList.Add(walkingZRow);
                        break;
                    case ContiType.PowerShort:
                        var powerShortRow = new ContiPowerShort(dcTestContiRow);
                        contiTestList.Add(powerShortRow);
                        break;
                    case ContiType.PowerSense:
                        var powerSenseRow = new ContiPowerSense(dcTestContiRow);
                        contiTestList.Add(powerSenseRow);
                        break;
                    case ContiType.GroundSense:
                        var groundSenseRow = new ContiGroundSense(dcTestContiRow);
                        contiTestList.Add(groundSenseRow);
                        break;
                    case ContiType.PowerImpedance:
                        var contiPowerSenseImpedance = new ContiSenseImpedance(dcTestContiRow);
                        contiTestList.Add(contiPowerSenseImpedance);
                        break;
                    case ContiType.GroundImpedance:
                        var groundSenseImpedance = new ContiSenseImpedance(dcTestContiRow);
                        contiTestList.Add(groundSenseImpedance);
                        break;

                    case ContiType.UserFunction:
                        var userFunction = new ContiUserFunction(dcTestContiRow);
                        contiTestList.Add(userFunction);
                        break;
                    case ContiType.Opcode:
                        var opcode = new ContiOpcode(dcTestContiRow);
                        contiTestList.Add(opcode);
                        break;
                    case ContiType.AutoZTtr:
                        var setAutoZTtr = new ContiAutoZTtr(dcTestContiRow);
                        contiTestList.Add(setAutoZTtr);
                        break;
                }
            }

            if (!contiTestList.Any(x =>
                string.Equals(
                    x.DcTestContiRow?.Category,
                    "AutoZTtr",
                    StringComparison.OrdinalIgnoreCase)))
            {
                var autoZTtr = new ContiAutoZTtr(new DcTestContiRow());
                contiTestList.Add(autoZTtr);
            }

            return contiTestList;
        }

        protected virtual SubFlowSheet GenerateContiFlow(List<ContiBase> contiTestList)
        {
            string postFixChiplet = "";
            string postFixSyntax = "";
            var flexJobs = new List<string>();

            if (TestPlanStatic.JobInfoSheet != null)
            {
                flexJobs.AddRange(TestPlanStatic.JobInfoSheet.Rows.Where(x => x.TesterType.ToUpper().Equals("UF")).Select(x => x.JobName));
            }
            else if (TestPlanStatic.Equipments.First().Equals(EnumEquipment.UltraFlex))
            {
                flexJobs.Add("All");
            }

            if (!string.IsNullOrEmpty(DcTestContiSheet.PostFixed))
            {
                postFixChiplet = DcTestContiSheet.PostFixed;
            }

            if (!string.IsNullOrEmpty(postFixChiplet))
            {
                postFixSyntax += "_" + postFixChiplet;
            }

            var contiFlow = new SubFlowSheet(ContiResult.FlowSheetName.TrimEnd('_') + postFixSyntax, DcTestContiSheet.SheetName);

            try
            {
                var status = new RelayStatus();

                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.Print, $"\"{contiFlow.Name} Start\""));
                contiTestList
                    .SelectMany(x => x.ContiFailFlags.Values)
                    .Distinct()
                    .ToList().ForEach(x => contiFlow.AddRow(CreateSimpleFlowRow(OpCode.FlagClear, x)));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.FlagClear, DcContiConst.FlagNamePowerShort));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.FlagClear, DcContiConst.FlagNamePowerSense));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.FlagClear, DcContiConst.FlagNameGroundSense));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.FlagClear, DcContiConst.FlagNameCres));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.FlagClear, DcContiConst.FPowerUpAlarm));
                FlagInitFromUfDefine(contiTestList, ref contiFlow);
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.Test, "SetPPMU_Clamp_Conti"));
                if (_needSetPowerAlarm)
                {
                    contiFlow.AddRow(CreateSimpleFlowRow(OpCode.Test, "SetPOWER_Alarm"));
                }
                //Added AutoZ TTR Instance
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.Test, contiTestList.FirstOrDefault(x => x.DcTestContiRow.TestType.Equals(ContiType.AutoZTtr))
                    ?.DcTestContiRow.InstanceName
                    ?? "DC_Continuity_IONeg_AUTOZ", "AutoZOnly"));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.BinTable, DcContiConst.BinNameOpenShort, "AutoZOnly"));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.BinTable, DcContiConst.BinNameOpen, "AutoZOnly"));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.BinTable, DcContiConst.BinNameShort, "AutoZOnly"));
                contiFlow.Rows.AddRange(CreateTryZFlow());

                HandlePPMUFlag(contiTestList, contiFlow);
                TestProgram.IgxlWorkBk.GenEnableToMainInitEnableWd("ForcePPMU", OpCode.Nop, "", FolderStructure.DirMain, "");

                //Generate DC Conti Core Body
                foreach (ContiBase test in contiTestList)
                {
                    List<FlowRow> flowRows = test.GenerateFlowRows();
                    if (flowRows == null)
                    {
                        continue;
                    }


                    RemoveBinOutFlowRowsIfNeeded(
                          flowRows,
                          test.DcTestContiRow.DisableBinOut
                          );


                    var onlyTestFlowRows = flowRows.Where(x => x.Opcode.Equals("Test", StringComparison.OrdinalIgnoreCase)).ToList();
                    if (!string.IsNullOrEmpty(test.DcTestContiRow.FailFlag))
                    {
                        onlyTestFlowRows.ForEach(x => x.FailAction = test.DcTestContiRow.FailFlag);
                    }
                    if (!string.IsNullOrEmpty(test.DcTestContiRow.SiteFlag))
                    {
                        List<string> siteFlag = test.DcTestContiRow.SiteFlag.Split(':').ToList();
                        if (siteFlag.Count > 1)
                        {
                            onlyTestFlowRows.ForEach(x => x.DeviceCondition = siteFlag[0].Trim());
                            onlyTestFlowRows.ForEach(x => x.DeviceName = siteFlag[1].Trim());
                        }
                        else
                        {
                            string condition = test.DcTestContiRow.SiteFlag.Contains("!") ? OpCode.FlagFalse : OpCode.FlagTrue;
                            onlyTestFlowRows.ForEach(x => x.DeviceCondition = condition);
                            onlyTestFlowRows.ForEach(x => x.DeviceName = test.DcTestContiRow.SiteFlag.Replace("!", ""));
                        }
                    }
                    contiFlow.Rows.AddRange(flowRows);
                }

                //Insert TryZ Set Device
                try
                {
                    int contiLatestIndex = contiFlow.Rows.Select((value, index) => new { Value = value, Index = index })
                        .Where(p => p.Value.Parameter.Contains("Bin_DC_short"))
                        .Select(pair => pair.Index).ToList().Last();
                    contiFlow.Rows.InsertRange(contiLatestIndex + 1, CreateTryZFlow());
                }
                catch
                {
                    int contiLatestIndex = contiFlow.Rows.Select((value, index) => new { Value = value, Index = index })
                        .Select(pair => pair.Index).ToList().Last();
                    contiFlow.Rows.InsertRange(contiLatestIndex + 1, CreateTryZFlow());
                }


                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.BinTable, DcContiConst.BinNamePowerSense));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.BinTable, DcContiConst.BinNameGndSense));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.BinTable, DcContiConst.BinNameCres));

                if (NwireSingleton.Instance().HasNwirePin && flexJobs.Any())
                {
                    FlowRow sbClockRow = CreateSimpleFlowRow("Test", "StartSBClock");
                    if (!flexJobs.Exists(x => x.Equals("All")))
                    {
                        sbClockRow.Job = string.Join(",", flexJobs);
                    }

                    contiFlow.AddRow(sbClockRow);
                }

                if (_needPowerUp)
                {
                    if (TestPlanStatic.UfInstanceTable == null || (TestPlanStatic.UfInstanceTable != null && !TestPlanStatic.UfInstanceTable.CheckPowerUpExist()))
                    {
                        contiFlow.AddRow(CreateSimpleFlowRow(OpCode.Test, ("PowerUp" + "_").Trim('_'), "", "F_PowerUp_Alarm"));
                    }
                    contiFlow.AddRow(CreateSimpleFlowRow(OpCode.BinTable, DcContiConst.BinPowerUpAlarm));
                }

                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.Print, $"\"{contiFlow.Name} End\""));
                contiFlow.AddRow(CreateSimpleFlowRow(OpCode.Return, ""));

            }
            catch (Exception)
            {
                string errMsg = "The TestPlan Error.";
                Response.Report(errMsg, EnumMessageLevel.Error, 0);
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_01, NeededSheets.ContiDcTest, 1, 1, "The TestPlan Error.");
            }


            return contiFlow;

        }


        internal void HandlePPMUFlag(
            List<ContiBase> contiTestList,
            SubFlowSheet contiFlow)
        {
            List<string> flags = GetAllContinuityFlags(contiTestList);
            string flagStr = BuildFlagString(flags);

            if (!string.IsNullOrEmpty(flagStr))
            {
                contiFlow.AddRow(
                    this.CreateSimpleFlowRow(
                        OpCode.FlagTrue, flagStr, "ForcePPMU"));
            }
        }


        internal static List<string> GetAllContinuityFlags(List<ContiBase> contiTestList)
        {
            return contiTestList
                .FindAll(x => Regex.IsMatch(
                    x.DcTestContiRow.Category,
                    "Continuity",
                    RegexOptions.IgnoreCase))
                .Where(x => !string.IsNullOrEmpty(x.DcTestContiRow.SiteFlag))
                .Select(x => x.DcTestContiRow.SiteFlag)
                .Distinct()
                .ToList();
        }

        internal static string BuildFlagString(List<string> flags)
        {
            if (!flags.Any())
            {
                return null;
            }
            return flags.Count == 1
                ? flags[0]
                : string.Join(", ", flags);
        }


        public static int RemoveBinOutFlowRowsIfNeeded(
            List<FlowRow> flowRows,
            string disableBinOut)
        {
            if (flowRows == null)
            {
                return 0;
            }

            if (!string.Equals(disableBinOut, "true", StringComparison.OrdinalIgnoreCase))
            {
                return 0;
            }

            const string keyword = "Bin";

            return flowRows.RemoveAll(r =>
                r != null && (
                    (!string.IsNullOrEmpty(r.Opcode) &&
                     r.Opcode.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (!string.IsNullOrEmpty(r.Parameter) &&
                     r.Parameter.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                )
            );
        }


        private void FlagInitFromUfDefine(List<ContiBase> contiTestList, ref SubFlowSheet contiFlow)
        {
            IEnumerable<List<string>> allFailFlags = contiTestList.Where(x => !string.IsNullOrEmpty(x.DcTestContiRow.FailFlag)).Select(x => x.DcTestContiRow.FailFlag.Replace(" ", "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList());
            var addedFlags = new List<string>();
            foreach (List<string> flags in allFailFlags)
            {
                foreach (string flag in flags)
                {
                    if (addedFlags.Exists(x => x.Equals(flag, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    contiFlow.AddRow(CreateSimpleFlowRow(OpCode.FlagClear, flag));
                    addedFlags.Add(flag);
                }
            }
        }

        protected virtual InstanceSheet GenerateContiInstanceSheet(List<ContiBase> contiTestList)
        {
            string postFixChiplet = "";
            string postFixSyntax = "";
            if (!string.IsNullOrEmpty(DcTestContiSheet.PostFixed))
            {
                postFixChiplet = DcTestContiSheet.PostFixed;
            }

            if (!string.IsNullOrEmpty(postFixChiplet))
            {
                postFixSyntax += "_" + postFixChiplet;
            }

            var contiInstance = new InstanceSheet(ContiResult.InstanceSheetName.TrimEnd('_') + postFixSyntax, DcTestContiSheet.SheetName);
            try
            {
                foreach (ContiBase test in contiTestList)
                {
                    List<InstanceRow> instRows = test.GenerateInstanceRows();
                    contiInstance.Rows.AddRange(instRows);
                }
            }
            catch (Exception)
            {
                string errMsg = "The TestPlan Error.";
                Response.Report(errMsg, EnumMessageLevel.Error, 0);
                ErrorReportManager.AddError(BasicErrorType.E_FormatError_02, NeededSheets.ContiDcTest, 1, 1, "The TestPlan Error.");
            }

            return contiInstance;
        }

        protected virtual List<InstanceRow> GenerateCommonInstanceRows(List<ContiBase> contiTestList, string subblock = "")
        {
            var commonInstanceRows = new List<InstanceRow>();

            //Add Power Up 
            if (_needPowerUp)
            {
                if (TestPlanStatic.UfInstanceTable == null || (TestPlanStatic.UfInstanceTable != null && !TestPlanStatic.UfInstanceTable.CheckPowerUpExist()))
                {
                    commonInstanceRows.Add(GenPowerUpDownInstanceRow("PowerUp", subblock));
                }
            }

            GenSetPpmuClampContiInstanceRow(contiTestList, commonInstanceRows);
            return commonInstanceRows;
        }

        internal void GetClampSettingFromContiPlanByUser(List<ContiBase> contiTestList, ref Dictionary<double, List<string>> clampLimit)
        {

            string condition = contiTestList
                .FirstOrDefault(x => string.Equals(x.DcTestContiRow?.InstanceName, "SetPPMU_Clamp_Conti", StringComparison.OrdinalIgnoreCase))
                ?.DcTestContiRow?.Condition;

            Dictionary<string, string> conditions = condition?
                .Split(';', StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Split('=', 2))
                .Where(x => x.Length == 2)
                .ToDictionary(
                    x => x[0],
                    x => x[1],
                    StringComparer.OrdinalIgnoreCase)
                   ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            ConfigureClampLimits(conditions, clampLimit, "clampHiGroup", "vch");
            ConfigureClampLimits(conditions, clampLimit, "clampLoGroup", "vcl");
        }


        internal void ConfigureClampLimits(
            Dictionary<string, string> conditions,
            Dictionary<double, List<string>> result,
            string nameKey,
            string valueKey)
        {
            if (!conditions.TryGetValue(nameKey, out string namesStr) ||
                !conditions.TryGetValue(valueKey, out string valuesStr))
            {
                return;
            }

            var names = namesStr.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(x => x.Trim())
                                .ToList();

            var values = valuesStr.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                  .Select(x => double.Parse(x.Trim()))
                                  .ToList();

            for (int i = 0; i < Math.Min(names.Count, values.Count); i++)
            {
                if (!result.ContainsKey(values[i]))
                {
                    result[values[i]] = new List<string>();
                }

                result[values[i]].Add(names[i]);
            }
        }


        private void GetClampSettingFromIoInfo(List<IoInfoRow> ioInfoRows, ref Dictionary<double, List<string>> clampLimit)
        {
            foreach (IoInfoRow ioInfoRow in ioInfoRows.Where(x => !string.IsNullOrEmpty(x.Vch) || !string.IsNullOrEmpty(x.Vcl)))
            {
                if (double.TryParse(ioInfoRow.Vch, out double value))
                {
                    if (!clampLimit.ContainsKey(value))
                    {
                        clampLimit[value] = new List<string>();
                    }
                    clampLimit[value].Add(ioInfoRow.PinGrpName);
                }
                if (double.TryParse(ioInfoRow.Vcl, out value))
                {
                    if (!clampLimit.ContainsKey(value))
                    {
                        clampLimit[value] = new List<string>();
                    }
                    clampLimit[value].Add(ioInfoRow.PinGrpName);
                }
            }
        }

        private void GetClampSettingFromContiPlan(List<ContiBase> contiTestList, ref Dictionary<double, List<string>> clampLimit)
        {
            var contiSetClampTestRows = new List<ContiBase>();
            foreach (ContiBase contiRow in contiTestList)
            {
                List<string> pinList = contiRow.TestPinGroup.Split(';').ToList();

                foreach (string pin in pinList)
                {
                    if (TestProgram.IgxlWorkBk.PinMapPair.Value.GetGroup(pin) != null && TestProgram.IgxlWorkBk.PinMapPair.Value.GetGroup(pin).PinType == PinMapConst.TypeIo)
                    {
                        contiSetClampTestRows.Add(contiRow);
                        break;
                    }
                }

            }

            foreach (ContiBase setClampRow in contiSetClampTestRows)
            {
                var hiLimitValues = new List<double>();
                foreach (DcTestContiSheetLimit limit in setClampRow.DcTestContiRow.Limits)
                {
                    if (double.TryParse(limit.HiLimitValue, out double tryValue))
                    {
                        hiLimitValues.Add(tryValue);
                    }
                }
                if (!hiLimitValues.Any())
                {
                    continue;
                }

                double setClamp = 0.0;
                if (hiLimitValues.Max() < 0)
                {
                    setClamp = -1.2;
                }
                else if (hiLimitValues.Max() <= 1.2)
                {
                    setClamp = 1.2;
                }
                else if (hiLimitValues.Max() > 1.2)
                {
                    setClamp = 1.2 + 0.3;
                }

                if (setClampRow is ContiSenseImpedance)
                {
                    setClamp = 1.2;
                }

                IEnumerable<string> clampPinGroup = setClampRow.TestPinGroup
                                                    .ToUpper()
                                                    .Split(';')
                                                    .Where(x => TestProgram.IgxlWorkBk.PinMapPair.Value.GetGroup(x) != null
                                                                && TestProgram.IgxlWorkBk.PinMapPair.Value.GetGroup(x).PinType == PinMapConst.TypeIo);

                if (clampLimit.ContainsKey(setClamp))
                {
                    clampLimit[setClamp].AddRange(clampPinGroup);
                }
                else
                {
                    clampLimit[setClamp] = clampPinGroup.ToList();
                }
            }
        }

        protected virtual void GenSetPpmuClampContiInstanceRow(List<ContiBase> contiTestList, List<InstanceRow> commonInstanceRows)
        {
            InstanceRow row = new InstanceRow { TestName = "SetPPMU_Clamp_Conti" };

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CsFuncNameSetPpmuClamp, "conti", true);
            if (!function.IsFound)
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameSetPpmuClamp, "conti");
            }

            row.VbtName = function.FullFunctionName;
            row.VbtType = function.Type;

            var clampLimit = new Dictionary<double, List<string>>();
            if (_ioInfoSheet.BlockIoInfo.TryGetValue("DCTEST_Continuity", out List<IoInfoRow> value))
            {
                GetClampSettingFromIoInfo(value, ref clampLimit);
            }
            else
            {
                GetClampSettingFromContiPlan(contiTestList, ref clampLimit);
            }

            if (contiTestList.Any(x => x.DcTestContiRow.InstanceName.Equals("SetPPMU_Clamp_Conti")))
            {
                clampLimit.Clear();
                GetClampSettingFromContiPlanByUser(contiTestList, ref clampLimit);
            }

            var clampSettings = clampLimit.ToList();
            if (function.Type.ToUpper() == ".NET")
            {
                var posClampPin = new List<string>();
                var posClampValue = new List<string>();

                clampSettings.Where(x => x.Key >= 0).ToList().ForEach(x => x.Value.ForEach(y =>
                {
                    posClampPin.Add(y);
                    posClampValue.Add(x.Key.ToString(CultureInfo.InvariantCulture));
                }));
                function.SetParamValue("clampHiGroup", string.Join(",", posClampPin));
                function.SetParamValue("vch", string.Join(",", posClampValue));

                var negClampPin = new List<string>();
                var negClampValue = new List<string>();

                clampSettings.Where(x => x.Key < 0).ToList().ForEach(x => x.Value.ForEach(y =>
                {
                    negClampPin.Add(y);
                    negClampValue.Add(x.Key.ToString(CultureInfo.InvariantCulture));
                }));
                if (negClampPin.Count > 0 &&
                    negClampValue.Count > 0 &&
                    negClampPin.Count == negClampValue.Count)

                {
                    function.SetParamValue("clampLoGroup", string.Join(",", negClampPin));
                    function.SetParamValue("vcl", string.Join(",", negClampValue));
                }
            }
            else
            {
                var posClamp = clampSettings.Where(x => x.Key >= 0).ToList();
                for (int i = 1; i <= 7; i++)
                {
                    if (posClamp.Count < i)
                    {
                        break;
                    }

                    function.SetParamValue($"Pin_GP{i}", string.Join(",", posClamp[i - 1].Value.ToList()));
                    function.SetParamValue($"Pin_GP{i}_Vch", posClamp[i - 1].Key.ToString(CultureInfo.InvariantCulture));
                }
                KeyValuePair<double, List<string>> negClamp = clampSettings.FirstOrDefault(x => x.Key == -1.2);
                if (negClamp.Value != null)
                {
                    function.SetParamValue("Pin_GP", string.Join(",", negClamp.Value));
                    function.SetParamValue("Pin_GP_Vcl", negClamp.Key.ToString(CultureInfo.InvariantCulture));
                }
            }
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;
            commonInstanceRows.Add(row);
        }

        protected virtual InstanceRow GenPowerUpDownInstanceRow(string testName, string subblock = "")
        {
            var lRow = new InstanceRow
            {
                TestName = (testName + "_" + subblock).TrimEnd('_'),
                DcCategory = NwireSingleton.Instance().HasNwirePin ? GetNwireDcCategoryName() : "",
                DcSelector = NwireSingleton.Instance().HasNwirePin ? "Typ" : ""
            };

            bool hasUltraFLex = TestPlanStatic.Equipments.Contains(EnumEquipment.UltraFlex);
            if (hasUltraFLex)
            {
                lRow.AcCategory = NwireSingleton.Instance().HasNwirePin ? BasicInitial.Common : "";
                lRow.AcSelector = NwireSingleton.Instance().HasNwirePin ? "Typ" : "";
                lRow.TimeSets = NwireSingleton.Instance().HasNwirePin ? "TimeSet_nWire" : "";
            }

            lRow.PinLevels = NwireSingleton.Instance().HasNwirePin ? "Levels_nWire" : "";

            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(testName == "PowerUp" ? VbtFunctionLibShared.PowerUp : VbtFunctionLibShared.PowerDown, "conti");
            lRow.VbtType = function.Type;
            if (function.Type.ToUpper() == ".NET")
            {
                function.SetParamValue("rampWaitTime", "0.001");
                function.SetParamValue("sequenceWaitTime", "0.001");
                function.SetParamValue("disconnectPins", (DcContiConst.PinGroupAllDigitalPowerUp + "_" + subblock).TrimEnd('_'));
            }
            else
            {
                function.SetParamValue("DisconnectPinList", (DcContiConst.PinGroupAllDigitalPowerUp + "_" + subblock).TrimEnd('_'));

                if (LocalSpecs.ExistPowerPinListDcvs)
                {
                    function.SetParamValue("PowerPinList_DCVS", (CommonConst.DcvsPower + "_" + subblock).TrimEnd('_'));
                }

                if (LocalSpecs.ExistPowerPinListDcvi)
                {
                    function.SetParamValue("PowerPinList_DCVI", (CommonConst.DcviPower + "_" + subblock).TrimEnd('_'));
                }

                if (LocalSpecs.ExistIoSeqHiLo)
                {
                    function.SetParamValue("IO_H", CommonConst.InitHiPins);
                    function.SetParamValue("IO_L", CommonConst.InitLoPins);
                }

                function.SetParamValue("WaitConnectTime", "0.001");
                function.SetParamValue("DebugFlag", "-1");
            }

            lRow.ArgList = function.Parameters;
            lRow.Args = function.ArgList;
            lRow.VbtName = function.FullFunctionName;
            return lRow;
        }

        protected string GetNwireDcCategoryName()
        {
            string dcCategory = MultiTestSettingSheetsSingleton.Instance().DcCategoryInfos.FindNwireCategory(out EnumMessageLevel msgLevel, out string errorMsg);
            if (!string.IsNullOrEmpty(errorMsg))
            {
                //Report error msg
                Response.Report(errorMsg, msgLevel, 90);
            }
            return dcCategory;
        }

        private List<FlowRow> CreateTryZFlow()
        {
            var resultRows = new List<FlowRow>();

            var flowRowSetDevice = new FlowRow
            {
                Enable = "AutoZOnly",
                Opcode = OpCode.SetDevice,
                BinFail = "9",
                SortFail = "9997",
                Result = "Fail-Stop"
            };

            resultRows.Add(flowRowSetDevice);

            return resultRows;
        }

        protected List<BinTableRow> GenerateBinTableRows(ref List<BinTableRow> binTableRows, List<ContiBase> contiTestList = null)
        {
            GenUserDefinedFailFlagBintable(contiTestList, binTableRows);

            GenContiBinTable(binTableRows, DcContiConst.BinNameOpen, new List<string> { DcContiConst.FlagNameOpen }, "DC", "open");

            GenContiBinTable(binTableRows, DcContiConst.BinNameShort, new List<string> { DcContiConst.FlagNameShort }, "DC", "short");

            GenContiBinTable(binTableRows, DcContiConst.BinNameOpenShort, new List<string> { DcContiConst.FlagNameOpen, DcContiConst.FlagNameShort }, "DC", "openshort");

            GenContiBinTable(binTableRows, DcContiConst.BinNamePowerShort, new List<string> { DcContiConst.FlagNamePowerShort }, "DC", "powershort");

            GenContiBinTable(binTableRows, DcContiConst.BinNamePowerOpen, new List<string> { DcContiConst.FlagNamePowerOpen }, "DC", "poweropen");

            GenContiBinTable(binTableRows, DcContiConst.BinNamePowerSense, new List<string> { DcContiConst.FlagNamePowerSense }, "DC", "powersense");

            GenContiBinTable(binTableRows, DcContiConst.BinNameGndSense, new List<string> { DcContiConst.FlagNameGroundSense }, "DC", "groundsense");

            GenContiBinTable(binTableRows, DcContiConst.BinNameAutoZCheck, new List<string> { DcContiConst.FlagNameAutoZ }, "DC", "AutoZ", "check");

            GenContiBinTable(binTableRows, DcContiConst.BinPowerUpAlarm, new List<string> { DcContiConst.FPowerUpAlarm }, "DC", "PowerUp", "Alarm");

            GenContiBinTable(binTableRows, DcContiConst.BinNameCres, new List<string> { DcContiConst.FlagNameCres }, "DC", "CRES");

            if (contiTestList != null)
            {
                GenUfBinTable(contiTestList, binTableRows);
            }

            return binTableRows;
        }

        private void GenContiBinTable(
            List<BinTableRow> binTableRows,
            string binName,
            List<string> failFlags,
            string module = "DC",
            string category1 = "",
            string category2 = "")
        {
            var binTableRow = new BinTableRow
            {
                Name = binName,
                ItemList = string.Join(",", failFlags),
                Op = "AND"
            };
            binTableRow.Items.AddRange(Enumerable.Repeat("T", failFlags.Count()));
            BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo(module, category1, category2, binTableRow);
            binTableRow.Sort = binNumInfo.SoftBin.ToString("G15");
            binTableRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
            binTableRow.Result = binNumInfo.BinNumInfo.Status;
            if (!binTableRows.Exists(x => x.GetUniqKey().Equals(binTableRow.GetUniqKey())))
            {
                binTableRows.Add(binTableRow);
            }
        }

        private static void GenUfBinTable(List<ContiBase> contiTestList, List<BinTableRow> binTableRows)
        {
            var allUfRows = contiTestList.Where(x => x.DcTestContiRow.TestType.Equals(ContiType.UserFunction)).ToList();
            foreach (ContiBase row in allUfRows)
            {
                if (string.IsNullOrEmpty(row.DcTestContiRow.FailFlag))
                {
                    continue;
                }

                string[] flagList = row.DcTestContiRow.FailFlag.Replace(" ", "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                var binTableRow = new BinTableRow();
                string binName = "Bin_" + string.Join("_", flagList.Select(x => DcContiConst.FailFlagRegex.Replace(x.Trim(), "")));
                binTableRow.Name = binName;
                binTableRow.ItemList = string.Join(",", flagList);
                binTableRow.Op = "AND";
                binTableRow.Items = Enumerable.Repeat("T", flagList.Length).ToList();
                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("DC", "", "", binTableRow);
                binTableRow.Sort = binNumInfo.SoftBin.ToString("G15");
                binTableRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                binTableRow.Result = binNumInfo.BinNumInfo.Status;
                if (!binTableRows.Exists(x => x.GetUniqKey().Equals(binTableRow.GetUniqKey())))
                {
                    binTableRows.Add(binTableRow);
                }
            }
        }

        private void GenUserDefinedFailFlagBintable(List<ContiBase> contiTestList, List<BinTableRow> binTableRows)
        {
            var userDefinedFailFlagRows = contiTestList.Where(x => x.DcTestContiRow.ConditionDict.Values.Any()).ToList();
            foreach (ContiBase row in userDefinedFailFlagRows)
            {
                if (row.ContiFailFlags != null && row.ContiFailFlags.Count > 0)
                {
                    foreach (KeyValuePair<string, string> kv in row.ContiFailFlags)
                    {
                        var binTableRow = new BinTableRow();
                        string flagName = row.ContiFailFlags[kv.Key];
                        string[] nameSegs = DcContiConst.FailFlagRegex.Replace(flagName, "").Split('_');
                        binTableRow.Name = DcContiConst.FailFlagRegex.Replace(flagName, "Bin_DC_");
                        binTableRow.ItemList = row.ContiFailFlags[kv.Key];
                        binTableRow.Op = "AND";
                        binTableRow.Items.Add("T");
                        BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("DC", nameSegs[0], nameSegs.Length > 1 ? nameSegs[1] : "", binTableRow);
                        binTableRow.Sort = binNumInfo.SoftBin.ToString("G15");
                        binTableRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                        binTableRow.Result = binNumInfo.BinNumInfo.Status;

                        if (!binTableRows.Exists(x => x.GetUniqKey().Equals(binTableRow.GetUniqKey())))
                        {
                            binTableRows.Add(binTableRow);
                        }
                    }
                }
                if (!string.IsNullOrEmpty(row.CombinedBinName))
                {
                    var binTableRow = new BinTableRow();
                    string binName = row.CombinedBinName;
                    string[] nameSegs = DcContiConst.BinDcRegex.Replace(binName, "").Split('_');
                    binTableRow.Name = binName;
                    binTableRow.ItemList = string.Join(",", row.ContiFailFlags.Values);
                    binTableRow.Op = "AND";
                    binTableRow.Items = Enumerable.Repeat("T", row.ContiFailFlags.Values.Count).ToList();
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("DC", nameSegs.Length > 1 ? nameSegs[0] + nameSegs[1] : "", "", binTableRow);
                    binTableRow.Sort = binNumInfo.SoftBin.ToString("G15");
                    binTableRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                    binTableRow.Result = binNumInfo.BinNumInfo.Status;

                    if (!binTableRows.Exists(x => x.GetUniqKey().Equals(binTableRow.GetUniqKey())))
                    {
                        binTableRows.Add(binTableRow);
                    }
                }
            }
        }

        protected FlowRow CreateSimpleFlowRow(string opCode, string parameter, string enableWord = "", string failAction = "")
        {
            var row = new FlowRow { Opcode = opCode, Parameter = parameter, Enable = enableWord, FailAction = failAction };
            return row;
        }
    }
}
