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
    public class ContiGroundSense : ContiBase
    {
        private readonly string _powerPinsArg;
        private readonly string _digitalPinsArg;
        private readonly string _forceV = "0.1";
        private readonly string _chForceI = "0";

        public ContiGroundSense(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
        {
            List<string> pinGroup = DcTestContiRow.PinGroup.Split(';').ToList();
            if (pinGroup.Count == 1)
            {
                _powerPinsArg = pinGroup[0];
                _digitalPinsArg = pinGroup[0];

            }
            else
            {
                _powerPinsArg = pinGroup[0];
                _digitalPinsArg = pinGroup[1];
            }

            Dictionary<string, string> forceCondition = DcTestContiRow.GetForceConditions();
            KeyValuePair<string, string> forceV = forceCondition.FirstOrDefault(x => x.Key.StartsWith("V", StringComparison.OrdinalIgnoreCase));
            KeyValuePair<string, string> chForceI = forceCondition.FirstOrDefault(x => x.Key.StartsWith("I", StringComparison.OrdinalIgnoreCase));

            _forceV = forceV.Key == null ? _forceV : forceV.Value;
            _chForceI = chForceI.Key == null ? _chForceI : chForceI.Value;

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
                InstanceRow row = new InstanceRow();

                var allLowLimit = DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper()).Select(p => p.LimitValue).ToList();
                var allHiLimit = DcTestContiRow.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.HiLimitValue).ToList();

                row.TestName = DcTestContiRow.InstanceName + "_" + job;
                row.DcCategory = SetCategoryFromCondition();
                row.DcSelector = "Typ";
                row.PinLevels = GetLevels();

                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameSensePinConti, "conti", true);
                if (function.Type == ".NET")
                {
                    function.SetParamValue("powerPin", _powerPinsArg);
                    function.SetParamValue("sensePins", _digitalPinsArg);
                    function.SetParamValue("lowLimit", string.Join(",", allLowLimit));
                    function.SetParamValue("hiLimit", string.Join(",", allHiLimit));
                    function.SetParamValue("powerForceVoltage", _forceV);
                    function.SetParamValue("senseForceCurrent", _chForceI);
                    function.SetParamValue("isParallel", "True");
                    function.SetParamValue("settleTime", "0.005");
                }
                else
                {
                    function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameGndSense, "conti");
                    function.ArgList[0] = _powerPinsArg;
                    function.ArgList[1] = _digitalPinsArg;
                    function.ArgList[2] = string.Join(",", allLowLimit);
                    function.ArgList[3] = string.Join(",", allHiLimit);
                    function.ArgList[4] = ForceCondition == null ? _forceV : ForceCondition.ForceValue;
                    function.ArgList[5] = _chForceI;
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
    }
}
