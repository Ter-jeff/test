using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Reader;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public class ContiPowerShort : ContiBase
    {
        public ContiPowerShort(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
        {
            TestPinGroup = TestPinGroup.Split(';').FirstOrDefault();
        }

        public override List<FlowRow> GenerateFlowRows()
        {
            var flowRows = new List<FlowRow>();
            foreach (string job in DcTestContiRow.JobNameList)
            {
                string testName = DcTestContiRow.InstanceName + "_" + job;
                FlowRow row = CreateTestFlowRow(testName, DcContiConst.FlagNamePowerShort);
                string enable = $"{job}" + (string.IsNullOrEmpty(DcTestContiRow.EnableWord) ? "" : $"&&{DcTestContiRow.EnableWord}");
                row.Enable = enable;
                row.FailAction = DcContiConst.FlagNamePowerShort;

                flowRows.Add(row);
                //Gen use-limit
                var grpJobLimit = DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job).ToList();

                flowRows.AddRange(GenUseLimit(testName, grpJobLimit, enable));
            }

            List<string> result = GetBinFlags(flowRows, new[] { DcContiConst.FlagNamePowerShort, DcContiConst.FlagNamePowerOpen });
            if (result?.Any() == true)
            {
                ContiFailFlags[DcContiConst.FlagNamePowerShort] = "F_" + string.Join("_", result);
            }
            foreach (string flag in result)
            {
                flowRows.Add(CreateBinTableFlowRow($"Bin_{flag}"));
            }

            flowRows.Add(CreateBinTableFlowRow(DcContiConst.BinNamePowerShort));
            flowRows.Add(CreateBinTableFlowRow(DcContiConst.BinNamePowerOpen));

            return flowRows;
        }

        internal static List<string> GetBinFlags(
            IEnumerable<FlowRow> flowRows,
            IEnumerable<string> removeFlags)
        {
            var removeList = new HashSet<string>(
                removeFlags,
                StringComparer.OrdinalIgnoreCase);

            return flowRows
                .Where(x => !string.IsNullOrWhiteSpace(x.FailAction))
                .SelectMany(x => x.FailAction.Split(',', StringSplitOptions.RemoveEmptyEntries))
                .Select(x => x.Trim())
                .Where(x => !removeList.Contains(x))
                .Select(x => x.StartsWith("F_", StringComparison.OrdinalIgnoreCase)
                    ? x.Substring(2)
                    : x)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {
            var resultInstanceRows = new List<InstanceRow>();
            resultInstanceRows.AddRange(InstanceRowBase(DcTestContiRow));
            return resultInstanceRows;
        }

        private List<InstanceRow> InstanceRowBase(DcTestContiRow item)
        {
            List<InstanceRow> results = new List<InstanceRow>();
            foreach (string job in item.JobNameList)
            {
                var row = new InstanceRow();
                string testName = item.InstanceName + "_" + job;
                row.TestName = testName;
                row.DcCategory = SetCategoryFromCondition();
                row.DcSelector = "Typ";
                row.PinLevels = GetLevels();

                var allForceConditions = item.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.ForceConditionValue).ToList();
                item.GetForceConditions(string.Join(";", allForceConditions));

                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNamePowerShort, "conti", true);
                if (function.Type == ".NET")
                {
                    GenerateCSharpInstanceRow(ref function, row, item, job);
                }
                else
                {
                    if (Regex.IsMatch(item.Category, "serial", RegexOptions.IgnoreCase))
                    {
                        function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNamePowerShort, "conti");
                    }
                    else
                    {
                        function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.P2PShortPowerFvmiParallel, "conti");
                    }
                    GenerateVbtInstanceRow(ref function, item, job);
                }
                SetArgumentFromCondition(ref function);

                row.VbtName = function.FullFunctionName;
                row.VbtType = function.Type;
                row.ArgList = function.Parameters;
                row.Args = function.ArgList;
                results.Add(row);
            }
            return results;
        }

        public void GenerateCSharpInstanceRow(ref Function function, InstanceRow row, DcTestContiRow item, string job)
        {
            row.VbtType = ".NET";
            row.VbtName = function.FullFunctionName;

            //Default VBT setting
            function.SetParamValue("testLimitMode", "3");
            function.SetParamValue("initialCRFromLimit", "TRUE");
            function.SetParamValue("disconnectDigitalPins", "All_Digital");

            if (Regex.IsMatch(item.Category, "parallel", RegexOptions.IgnoreCase))
            {
                function.SetParamValue("maxForceVoltage", "3.5");
                function.SetParamValue("stepVoltage", "0.005");
                function.SetParamValue("initialCRFromLimit", "-1");
            }

            function.SetParamValue("measuredPins", string.Join(",",
                item.Limits.Where(p => p.LimitStage.Equals(job, StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(p.LimitHeader))
                .Select(p => p.LimitHeader)
                .ToList()));
            var allForceConditions = item.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.ForceConditionValue).ToList();
            function.SetParamValue("forceCondition", string.Join(",", item.GetForceConditions(string.Join(",", allForceConditions))));
            function.SetParamValue("lowLimit", "");
            function.SetParamValue("highLimit", "");
        }

        public void GenerateVbtInstanceRow(ref Function function, DcTestContiRow item, string job)
        {
            if (Regex.IsMatch(item.Category, "serial", RegexOptions.IgnoreCase))
            {
                //Default VBT setting
                function.SetParamValue("TestLimitMode", "3");
                function.SetParamValue("FlowLimitForInitIRange", "TRUE");
                function.SetParamValue("digital_pins", "All_Digital");
            }
            else
            {
                function.SetParamValue("MaxForceV", "3.5");
                function.SetParamValue("PinGroup_Cnt", "20");
                function.SetParamValue("Step_Voltage_Level", "0.005");
                function.SetParamValue("TestLimitMode", "3");
                function.SetParamValue("FlowLimitForInitIRange", "-1");
                function.SetParamValue("digital_pins", "All_Digital");
            }

            PinGroup pinGrp = TestProgram.IgxlWorkBk.PinMapPair.Value.GroupList.FirstOrDefault(p => Regex.IsMatch(p.PinName, "^" + TestPinGroup, RegexOptions.IgnoreCase));
            if (pinGrp != null)
            {
                string pinList = string.Join(";", pinGrp.PinList.Select(p => p.PinName));
                function.SetParamValue("allPowerPins", pinList);
                var forceConditions = new List<string>();
                var allForceConditions = item.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.ForceConditionValue).ToList();
                for (int i = 0; i < pinGrp.PinList.Count; ++i)
                {
                    forceConditions.AddRange(allForceConditions);
                }

                function.SetParamValue("ForceCondition", string.Join(";", item.GetForceConditions(string.Join(";", forceConditions))));

            }
            else
            {
                function.SetParamValue("allPowerPins", string.Join(";", item.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper()).Select(p => p.LimitHeader).ToList()));
                var allForceConditions = item.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.ForceConditionValue).ToList();
                function.SetParamValue("ForceCondition", string.Join(";", item.GetForceConditions(string.Join(";", allForceConditions))));
            }

            function.SetParamValue("LowLimit", "");
            function.SetParamValue("HiLimit", "");
        }

        public List<FlowRow> GenUseLimit(string instanceName, List<DcTestContiSheetLimit> useLimitRows, string job)
        {
            var flowRows = new List<FlowRow>();

            string testName = instanceName;

            foreach (DcTestContiSheetLimit limitRow in useLimitRows.Where(x => !string.IsNullOrWhiteSpace(x?.LimitHeader)))
            {
                var limitPins = new List<string> { limitRow.LimitHeader };
                PinGroup pinGrp = TestProgram.IgxlWorkBk.PinMapPair.Value.GroupList.FirstOrDefault(p => p.PinName.Equals(limitRow.LimitHeader, StringComparison.OrdinalIgnoreCase));
                BuildLimitPins(pinGrp);
                foreach (string pinName in limitPins)
                {
                    var row = new FlowRow { Opcode = OpCode.UseLimit, Parameter = testName, TName = limitRow.LimitHeader };
                    string highUnit = "A";
                    string highScale = ScaleHelper.GetHighScale(limitRow.LimitUnit);
                    string lowUnit = "";
                    string lowScale = "";
                    row.LoLim = limitRow.LimitValue;
                    row.HiLim = limitRow.HiLimitValue;
                    row.Scale = lowScale == "" ? highScale : lowScale;
                    row.Units = lowUnit == "" ? highUnit : lowUnit;
                    row.FailAction = DcContiConst.FlagNamePowerShort + (string.IsNullOrEmpty(limitRow.FailFlag) ||
                        limitRow.FailFlag.Equals(DcContiConst.FlagNamePowerShort, StringComparison.OrdinalIgnoreCase)
                        ? ""
                        : $",{limitRow.FailFlag}");
                    row.Enable = job;
                    row.Job = limitRow.LimitStage;
                    row.Comment = "";
                    flowRows.Add(row);
                }
            }
            return flowRows;
        }
        internal static List<string> BuildLimitPins(PinGroup pinGrp)
        {
            var limitPins = new List<string>();

            if (pinGrp != null)
            {
                limitPins = pinGrp.PinList
                                  .Select(x => x.PinName)
                                  .ToList();
            }

            return limitPins;
        }

    }
}
