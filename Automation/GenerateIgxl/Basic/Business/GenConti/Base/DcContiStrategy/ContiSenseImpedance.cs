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
    public class ContiSenseImpedance : ContiBase

    {
        private readonly string _sinkPinGrp = "";
        private readonly string _measPinGrp = "";
        private readonly string _sinkI1 = "";
        private readonly string _sinkI2 = "";

        public ContiSenseImpedance(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
        {
            if (dcTestContiRow.TestType == ContiType.PowerImpedance)
            {
                _sinkI1 = "0.01";
                _sinkI2 = "0.05";
            }
            else if (dcTestContiRow.TestType == ContiType.GroundImpedance)
            {
                _sinkI1 = "0.000001";
                _sinkI2 = "0.02";
            }

            List<string> pinList = DcTestContiRow.PinGroup.Split(';').ToList();
            if (pinList.Count == 2)
            {
                _sinkPinGrp = pinList[0];
                _measPinGrp = pinList[1];
            }

            Dictionary<string, string> forceCondition = DcTestContiRow.GetForceConditions();
            if (forceCondition.Count == 2)
            {
                KeyValuePair<string, string> sink1 = forceCondition.FirstOrDefault(x => x.Key.Contains("1"));
                KeyValuePair<string, string> sink2 = forceCondition.FirstOrDefault(x => x.Key.Contains("2"));
                _sinkI1 = sink1.Key == null ? _sinkI1 : sink1.Value;
                _sinkI2 = sink2.Key == null ? _sinkI2 : sink2.Value;
            }
        }

        public override List<FlowRow> GenerateFlowRows()
        {
            List<FlowRow> rowList = new List<FlowRow>();

            foreach (string job in DcTestContiRow.JobNameList)
            {
                string instanceName = DcTestContiRow.InstanceName + "_" + job;
                FlowRow newRow = CreateTestFlowRow(instanceName, DcContiConst.FlagNamePowerSense);

                string enable = $"{job}" + (string.IsNullOrEmpty(DcTestContiRow.EnableWord) ? "" : $"&&{DcTestContiRow.EnableWord}");
                newRow.Enable = enable;
                rowList.Add(newRow);
            }
            return rowList;
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

                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameSenseImpedance, "conti", true);
                if (function.Type == ".NET")
                {
                    GenerateCSharpInstanceRow(ref function, job);
                }
                else
                {
                    function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameSenseImpedance, "conti");
                    GenerateVbtInstanceRow(ref function, job);
                }
                SetArgumentFromCondition(ref function);

                row.VbtName = function.FullFunctionName;
                row.VbtType = function.Type;
                row.ArgList = function.Parameters;
                row.Args = function.ArgList;
                resultInstanceRows.Add(row);
            }
            return resultInstanceRows;
        }

        public void GenerateCSharpInstanceRow(ref Function function, string job)
        {
            function.SetParamValue("sinkPins", string.Join(",", _sinkPinGrp));
            function.SetParamValue("measuredPins", string.Join(",", _measPinGrp));
            function.SetParamValue("sinkCurrent1", _sinkI1);
            function.SetParamValue("sinkCurrent2", _sinkI2);
            function.SetParamValue("measureRLowLimit", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper()).Select(p => p.LimitValue).ToList()));
            function.SetParamValue("measureRHighLimit", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.HiLimitValue).ToList()));
        }

        public void GenerateVbtInstanceRow(ref Function function, string job)
        {
            function.SetParamValue("Sink_Groups", string.Join(",", _sinkPinGrp));
            function.SetParamValue("Meas_Groups", string.Join(",", _measPinGrp));
            function.SetParamValue("Sink1_Current", _sinkI1);
            function.SetParamValue("Sink2_Current", _sinkI2);
            function.SetParamValue("LowLimit", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper()).Select(p => p.LimitValue).ToList()));
            function.SetParamValue("HiLimit", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.HiLimitValue).ToList()));
        }
    }
}
