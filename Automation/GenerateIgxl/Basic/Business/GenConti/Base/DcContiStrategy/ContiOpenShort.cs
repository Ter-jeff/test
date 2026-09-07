using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenLevel.Business;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;
using Automation.Reader;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public class ContiOpenShort : ContiBase
    {
        private readonly string _isinkIsource = "";

        public ContiOpenShort(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
        {
            ContiFailFlags = new Dictionary<string, string>
            {
                { DcContiConst.FlagNameOpen, DcContiConst.FlagNameOpen },
                { DcContiConst.FlagNameShort, DcContiConst.FlagNameShort },
            };
            if (dcTestContiRow.ConditionDict.TryGetValue("Flag_Open", out string flag))
            {
                ContiFailFlags[DcContiConst.FlagNameOpen] = flag;
            }
            if (dcTestContiRow.ConditionDict.TryGetValue("Flag_Short", out flag))
            {
                ContiFailFlags[DcContiConst.FlagNameShort] = flag;
            }
            if (dcTestContiRow.ConditionDict.TryGetValue("openFlagName", out flag))
            {
                ContiFailFlags[DcContiConst.FlagNameOpen] = flag;
            }
            if (dcTestContiRow.ConditionDict.TryGetValue("shortFlagName", out flag))
            {
                ContiFailFlags[DcContiConst.FlagNameShort] = flag;
            }
            CombinedBinName = "Bin_DC_" + string.Join("_", ContiFailFlags.Values.Select(x => DcContiConst.FailFlagRegex.Replace(x, "")));

            Dictionary<string, string> forceCondition = DcTestContiRow.GetForceConditions();
            KeyValuePair<string, string> isinkOrIsource = forceCondition.FirstOrDefault(x => x.Key.StartsWith("I", StringComparison.OrdinalIgnoreCase));

            _isinkIsource = isinkOrIsource.Key == null ? _isinkIsource : isinkOrIsource.Value;

        }
        public override List<FlowRow> GenerateFlowRows()
        {
            List<FlowRow> flowRows = new List<FlowRow>();

            foreach (string job in DcTestContiRow.JobNameList)
            {
                FlowRow ppmuRow = CreateTestFlowRow(DcTestContiRow.InstanceName + "_" + job, "");
                string enable = $"{job}" + (string.IsNullOrEmpty(DcTestContiRow.EnableWord) ? "" : $"&&{DcTestContiRow.EnableWord}");
                ppmuRow.Enable = enable;
                flowRows.Add(ppmuRow);
            }

            if (!string.IsNullOrEmpty(CombinedBinName))
            {
                flowRows.Add(CreateBinTableFlowRow(CombinedBinName));
            }
            foreach (string failFlag in ContiFailFlags.Values)
            {
                flowRows.Add(CreateBinTableFlowRow(DcContiConst.FailFlagRegex.Replace(failFlag, "Bin_DC_")));
            }

            return flowRows;
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {
            List<InstanceRow> resultInstanceRows = new List<InstanceRow>();
            foreach (string job in DcTestContiRow.JobNameList)
            {
                InstanceRow row = new InstanceRow();
                switch (LocalSpecs.Options.Device)
                {
                    case EnumDevice.AP:
                        row.TestName = DcTestContiRow.InstanceName + "_" + job;
                        row.DcCategory = SetCategoryFromCondition();
                        row.DcSelector = "Typ";
                        row.PinLevels = GetLevels();

                        Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameAutoZ, "conti", true);
                        if (function.Type == ".NET")
                        {
                            GenerateCSharpInstanceRow(ref function, job);
                        }
                        else
                        {
                            function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameAutoZ, "conti");
                            GenerateVbtInstanceRow(ref function, job);
                        }
                        SetArgumentFromCondition(ref function);

                        row.VbtName = function.FullFunctionName;
                        row.VbtType = function.Type;
                        row.ArgList = function.Parameters;
                        row.Args = function.ArgList;
                        resultInstanceRows.Add(row);
                        break;
                }
            }

            InstanceRow rowFunT = GenFunctionalT();

            resultInstanceRows.Add(rowFunT);
            return resultInstanceRows;

        }

        private InstanceRow GenFunctionalT()
        {
            InstanceRow rowFunT = new InstanceRow
            {
                TestName = CreateContiTestName(),
                DcCategory = SetCategoryFromCondition(),
                DcSelector = "Typ",
                TimeSets = TimeSetPlus.TsbContiPat,
                ColumnA = "Generated for Backup Usage"
            };
            if (ForceCondition.ForceValue.Contains("-"))
            {
                rowFunT.PinLevels = LevelInitial.WalkingZNegLevelSheetName;
            }
            else
            {
                rowFunT.PinLevels = LevelInitial.WalkingZPosLevelSheetName;
            }
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameFunctionalT, "conti", true);
            if (function.Type != ".NET")
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameFunctionalT, "conti");
            }

            rowFunT.VbtName = function.FullFunctionName;
            rowFunT.VbtType = function.Type;
            function.ArgList[0] = CreateWalkingZPatternNameOnly(true);
            if (LocalSpecs.Options.Device != EnumDevice.LCD)
            {
                function.ArgList[24] = "1";
            }

            rowFunT.ArgList = function.Parameters;
            rowFunT.Args = function.ArgList;
            return rowFunT;
        }

        public void GenerateCSharpInstanceRow(ref Function function, string job)
        {
            function.SetParamValue("measuredPins", string.Join(";", DcTestContiRow.Limits.Where(p => p.LimitStage == job).Select(p => p.LimitHeader).ToList()));
            function.SetParamValue("forceValue", _isinkIsource);
            function.SetParamValue("testLimitMode", "0");
            function.SetParamValue("disconnectPN", "TRUE");
            function.SetParamValue("connectPins", "All_Digital");
            function.SetParamValue("openFlagName", DcContiConst.FlagNameOpen);
            function.SetParamValue("shortFlagName", DcContiConst.FlagNameShort);
            function.SetParamValue("lowLimit", string.Join(";", DcTestContiRow.Limits.Where(p => p.LimitStage == job).Select(p => p.LimitValue)));
            function.SetParamValue("highLimit", string.Join(";", DcTestContiRow.Limits.Where(p => p.LimitStage == job).Select(p => p.HiLimitValue)));
        }

        public void GenerateVbtInstanceRow(ref Function function, string job)
        {
            function.SetParamValue("PinName", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage == job).Select(p => p.LimitHeader).ToList()));

            var allForceConditions = DcTestContiRow.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.ForceConditionValue).ToList();
            function.SetParamValue("ForceValue", string.Join(";", DcTestContiRow.GetForceConditions(string.Join(";", allForceConditions))));
            function.SetParamValue("TestLimitMode", "0");
            function.SetParamValue("PN_Disconnect", "-1");
            function.SetParamValue("connect_all_pins", "All_Digital_Disconnect_Continuity");
            function.SetParamValue("Flag_Open", DcContiConst.FlagNameOpen);
            function.SetParamValue("Flag_Short", DcContiConst.FlagNameShort);
            function.SetParamValue("Limit_L", string.Join(";", DcTestContiRow.Limits.Where(p => p.LimitStage == job).Select(p => p.LimitValue)));
            function.SetParamValue("Limit_H", string.Join(";", DcTestContiRow.Limits.Where(p => p.LimitStage == job).Select(p => p.HiLimitValue)));
        }
    }
}
