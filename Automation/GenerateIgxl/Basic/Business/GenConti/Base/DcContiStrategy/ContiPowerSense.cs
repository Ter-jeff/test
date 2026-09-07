using System;
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
    public class ContiPowerSense : ContiBase
    {
        private readonly string _powerPinsArg;
        private readonly string _sensePinAdditionNameArg;

        private readonly string _powerForceVoltage = "";
        private readonly string _senseForceCurrent = "0";

        public ContiPowerSense(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
        {
            List<string> list = DcTestContiRow.PinGroup.Split(';').ToList();
            if (list.Count == 1)
            {
                _powerPinsArg = list[0];
                _sensePinAdditionNameArg = list[0];

            }
            else
            {
                _powerPinsArg = list[0];
                _sensePinAdditionNameArg = list[1];
            }

            Dictionary<string, string> forceCondition = DcTestContiRow.GetForceConditions();
            KeyValuePair<string, string> forceV = forceCondition.FirstOrDefault(x => x.Key.StartsWith("V", StringComparison.OrdinalIgnoreCase));
            KeyValuePair<string, string> forceI = forceCondition.FirstOrDefault(x => x.Key.StartsWith("I", StringComparison.OrdinalIgnoreCase));

            _powerForceVoltage = forceV.Key == null ? _powerForceVoltage : forceV.Value;
            _senseForceCurrent = forceI.Key == null ? _senseForceCurrent : forceI.Value;
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


                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameSensePinConti, "conti", true);
                if (function.Type == ".NET")
                {
                    GenerateCSharpInstanceRow(ref function, job);
                }
                else
                {
                    function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNamePowerSense, "conti");
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
            if (!string.IsNullOrEmpty(_powerPinsArg) && !string.IsNullOrEmpty(_sensePinAdditionNameArg))
            {
                function.SetParamValue("powerPin", string.Join(",", _powerPinsArg));
                function.SetParamValue("sensePins", string.Join(",", _sensePinAdditionNameArg));
            }
            else
            {
                function.SetParamValue("powerPin", string.Join(",", TestPinGroup));
                function.SetParamValue("sensePins", string.Join(",", TestPinGroup));
            }
            function.SetParamValue("isParallel", "False");
            function.SetParamValue("powerForceVoltage", _powerForceVoltage);
            function.SetParamValue("lowLimit", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper()).Select(p => p.LimitValue).ToList()));//-2
            function.SetParamValue("hiLimit", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.HiLimitValue).ToList()));//20
            function.SetParamValue("senseForceCurrent", _senseForceCurrent);
            function.SetParamValue("settleTime", "0.05");
        }

        public void GenerateVbtInstanceRow(ref Function function, string job)
        {
            if (!string.IsNullOrEmpty(_powerPinsArg) && !string.IsNullOrEmpty(_sensePinAdditionNameArg))
            {
                function.SetParamValue("PowerPins", string.Join(",", _powerPinsArg));
                function.SetParamValue("SensePins", string.Join(",", _sensePinAdditionNameArg));
                function.SetParamValue("isUseParallel", "False");
            }
            else
            {
                function.SetParamValue("PowerPins", string.Join(",", TestPinGroup));
                function.SetParamValue("SensePins", string.Join(",", TestPinGroup));
                function.SetParamValue("isUseParallel", "False");
            }

            function.SetParamValue("Force_V", _powerForceVoltage);
            function.SetParamValue("LowLimit", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper()).Select(p => p.LimitValue).ToList()));//-2
            function.SetParamValue("HiLimit", string.Join(",", DcTestContiRow.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.HiLimitValue).ToList()));//20
        }
    }
}
