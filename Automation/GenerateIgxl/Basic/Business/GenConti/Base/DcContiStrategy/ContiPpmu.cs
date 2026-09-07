using System.Collections.Generic;
using System.Data;
using System.Linq;

using Automation.Const;
using Automation.Reader;
using Automation.Static;

using IgxlLib.IgxlBase;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public class ContiPpmu : ContiBase
    {
        public ContiPpmu(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
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
            CombinedBinName = "Bin_DC_" + string.Join("_", ContiFailFlags.Values.Select(x => DcContiConst.FailFlagRegex.Replace(x, "")));
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
                InstanceRow row = new InstanceRow
                {
                    TestName = DcTestContiRow.InstanceName + "_" + job,
                    DcCategory = SetCategoryFromCondition(),
                    DcSelector = "Typ",
                    PinLevels = GetLevels()
                };

                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameOpenShort, "conti");
                GenerateVbtInstanceRow(ref function, job);
                SetArgumentFromCondition(ref function);

                row.VbtName = function.FullFunctionName;
                row.VbtType = function.Type;
                row.ArgList = function.Parameters;
                row.Args = function.ArgList;
                resultInstanceRows.Add(row);
            }
            return resultInstanceRows;
        }

        public void GenerateVbtInstanceRow(ref Function function, string job)
        {
            function.SetParamValue("digital_pins", TestPinGroup);
            var allForceConditions = DcTestContiRow.Limits.Where(p => p.LimitStage == job.ToUpper())
                .Select(p => p.ForceConditionValue).ToList();
            function.SetParamValue("force_i",
                string.Join(";", DcTestContiRow.GetForceConditions(string.Join(";", allForceConditions))));
            function.SetParamValue("TestLimitMode", "0");
            function.SetParamValue("PN_Disconnect", "-1");
            function.SetParamValue("connect_all_pins", "All_Digital_Disconnect_Continuity");
            function.SetParamValue("Flag_Open", ContiFailFlags[DcContiConst.FlagNameOpen]);
            function.SetParamValue("Flag_Short", ContiFailFlags[DcContiConst.FlagNameShort]);
            List<string> loLimits = DcTestContiRow.Limits.Where(p => p.LimitStage == job).Select(p => p.LimitValue)
                .ToList();
            List<string> hiLimits = DcTestContiRow.Limits.Where(p => p.LimitStage == job).Select(p => p.HiLimitValue)
                .ToList();
            function.SetParamValue("LowLimit", loLimits.Any() ? loLimits[0] : "");
            function.SetParamValue("HiLimit", hiLimits.Any() ? hiLimits[0] : "");
            function.SetParamValue("LoLimit2", loLimits.Count > 1 ? loLimits[1] : "");
            function.SetParamValue("HiLimit2", hiLimits.Count > 1 ? hiLimits[1] : "");
        }
    }
}
