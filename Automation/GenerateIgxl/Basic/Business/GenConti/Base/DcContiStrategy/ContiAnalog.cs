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
    public class ContiAnalog : ContiBase
    {
        public ContiAnalog(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
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
            var flowRows = new List<FlowRow>();

            foreach (string job in DcTestContiRow.JobNameList)
            {
                string testName = DcTestContiRow.InstanceName + "_" + job;
                FlowRow row = CreateTestFlowRow(testName, "");
                string enable = $"{job}" + (string.IsNullOrEmpty(DcTestContiRow.EnableWord) ? "" : $"&&{DcTestContiRow.EnableWord}");
                row.Enable = enable;
                flowRows.Add(row);

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
                InstanceRow row = GenUVI80_Continuity(job);
                resultInstanceRows.Add(row);
            }

            return resultInstanceRows;
        }

        protected virtual InstanceRow GenUVI80_Continuity(string job)
        {
            InstanceRow row = new InstanceRow();

            IEnumerable<DcTestContiSheetLimit> allLimitFirst = DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper() && !string.IsNullOrEmpty(p.LimitHeader));
            IEnumerable<DcTestContiSheetLimit> allLimitSecond = DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper() && string.IsNullOrEmpty(p.LimitHeader));

            row.TestName = DcTestContiRow.InstanceName + "_" + job;
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.VbtFuncNameUvi80Continuity, "conti");

            row.DcCategory = SetCategoryFromCondition();
            row.DcSelector = "Typ";
            row.PinLevels = GetLevels();
            //digital_pins,force_i,LowLimit,HiLimit,TestLimitMode,PN_Disconnect,Separate_limit,
            function.SetParamValue("digital_pins", TestPinGroup);
            function.SetParamValue("force_i", ForceCondition.ForceValue);
            function.SetParamValue("TestLimitMode", "0");
            function.SetParamValue("PN_Disconnect", "-1");
            function.SetParamValue("connect_all_pins", "All_Digital_Disconnect_Continuity");
            function.SetParamValue("Flag_Open", ContiFailFlags[DcContiConst.FlagNameOpen]);
            function.SetParamValue("Flag_Short", ContiFailFlags[DcContiConst.FlagNameShort]);
            function.SetParamValue("LowLimit", string.Join(",", allLimitFirst.Select(x => x.LimitValue)));
            function.SetParamValue("HiLimit", string.Join(",", allLimitFirst.Select(x => x.HiLimitValue)));
            function.SetParamValue("LowLimit2", string.Join(",", allLimitSecond.Select(x => x.LimitValue)));
            function.SetParamValue("HiLimit2", string.Join(",", allLimitSecond.Select(x => x.HiLimitValue)));

            SetArgumentFromCondition(ref function);

            row.VbtName = function.FullFunctionName;
            row.VbtType = function.Type;
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;
            return row;
        }
    }
}
