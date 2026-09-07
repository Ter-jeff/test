using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;


namespace Automation.GenerateIgxl.Basic.Business.GenNwire.Business
{
    public class NwireFlow
    {
        protected const string Reset = "Reset";
        protected const string Enable = "Enable";
        protected const string Disable = "Disable";
        protected const string StartSupportBoard = "StartSBClock";
        protected const string FreeRunClkEnable = "FreeRunClk_Enable";
        protected const string FreeRunClkDisable = "FreeRunClk_Disable";

        public virtual List<SubFlowSheet> GenerateFlow()
        {
            var nWireFlows = new List<SubFlowSheet>();
            DataTable table = NwireSingleton.Instance().SettingInfo.SettingTable;
            Dictionary<string, string> dic = NwireSingleton.Instance().SettingInfo.PatternDic;
            var flexJobs = new List<string>();
            if (TestPlanStatic.JobInfoSheet != null)
            {
                flexJobs.AddRange(TestPlanStatic.JobInfoSheet.Rows.Where(x => x.TesterType.ToUpper().Equals("UF")).Select(x => x.JobName));
            }
            else if (TestPlanStatic.Equipments.First().Equals(EnumEquipment.UltraFlex))
            {
                flexJobs.Add("All");
            }

            for (int i = 0; i < table.Rows.Count; i++)
            {
                string item = table.Rows[i][0].ToString();
                SubFlowSheet flow = CreateSubFlow(item, dic);

                List<ProtocolAwarePin> nWirePin = NwireSingleton.Instance().SettingInfo.NwirePins;
                foreach (EnumEquipment testerType in TestPlanStatic.Equipments)
                {
                    foreach (ProtocolAwarePin awarePin in nWirePin)
                    {
                        awarePin.FlowControlAction = item;
                        if (Regex.IsMatch(table.Rows[i][awarePin.OutClk].ToString(), Reset, RegexOptions.IgnoreCase))
                        {
                            DisableOnePin(awarePin, flow, testerType);
                        }
                        else if (Regex.IsMatch(table.Rows[i][awarePin.OutClk].ToString(), Disable, RegexOptions.IgnoreCase))
                        {
                            DisableOnePin(awarePin, flow, testerType);
                        }
                    }
                }

                if (NwireSingleton.Instance().HasNwirePin && flexJobs.Any())
                {
                    var startSbc = new FlowRow { Opcode = OpCode.Test, Parameter = StartSupportBoard };
                    if (!flexJobs.Exists(x => x.Equals("All")))
                    {
                        startSbc.Job = string.Join(",", flexJobs);
                    }

                    flow.AddRow(startSbc);
                }

                foreach (EnumEquipment testerType in TestPlanStatic.Equipments)
                {
                    int awarePinCount = 1;
                    foreach (ProtocolAwarePin awarePin in nWirePin)
                    {

                        if (Regex.IsMatch(table.Rows[i][awarePin.OutClk].ToString(), Reset, RegexOptions.IgnoreCase))
                        {
                            EnableOnePin(awarePin, flow, testerType);
                        }
                        else if (Regex.IsMatch(table.Rows[i][awarePin.OutClk].ToString(), Enable, RegexOptions.IgnoreCase))
                        {
                            awarePin.ControlAction = table.Rows[i][awarePinCount].ToString();
                            EnableOnePin(awarePin, flow, testerType);

                            awarePinCount++;
                        }
                    }
                }
                var row = new FlowRow { Opcode = "Return" };
                flow.AddRow(row);
                nWireFlows.Add(flow);
            }

            return nWireFlows;
        }

        internal SubFlowSheet CreateSubFlow(
            string item,
            Dictionary<string, string> dic)
        {
            item = item?.Trim();

            if (string.IsNullOrEmpty(item))
            {
                return new SubFlowSheet(
                    NwireSetting.ConFlownWire.TrimEnd('_'),
                    "Autogen");
            }

            if (dic.TryGetValue(item, out string value))
            {
                var flow = new SubFlowSheet(
                    NwireSetting.ConFlownWire + value, "Autogen");
                PrintPattern(flow, item);
                return flow;
            }

            return new SubFlowSheet(
                NwireSetting.ConFlownWire + item, "Autogen");
        }

        internal void PrintPattern(SubFlowSheet flow, string pattern)
        {
            var row = new FlowRow { Parameter = "Flow_" + pattern + "_Strat", Opcode = "Print" };
            flow.AddRow(row);

        }
        protected void DisableOnePin(ProtocolAwarePin nWirePin, SubFlowSheet flow, EnumEquipment testerType)
        {
            var row = new FlowRow();
            string testType = testerType.Equals(EnumEquipment.UltraFlex) ? "UF" : "UFP";
            row.Parameter = FreeRunClkDisable + "_"
                + nWirePin.CreatePinNameWithDiff() + "_"
                + testType
                + (string.IsNullOrEmpty(nWirePin.FlowControlAction) ? "" : "_" + nWirePin.FlowControlAction);
            row.Opcode = OpCode.Test;
            if (TestPlanStatic.JobInfoSheet != null)
            {
                row.Job = string.Join(",", TestPlanStatic.JobInfoSheet.Rows.Where(x => x.TesterType.Equals(testType)).Select(x => x.JobName));
            }

            flow.AddRow(row);
        }

        protected void EnableOnePin(ProtocolAwarePin nWirePin, SubFlowSheet flow, EnumEquipment testerType)
        {
            var flowRow = new FlowRow();
            string testType = testerType.Equals(EnumEquipment.UltraFlex) ? "UF" : "UFP";
            flowRow.Parameter = FreeRunClkEnable + "_"
                + nWirePin.CreatePinNameWithDiff() + "_"
                + testType
                + (string.IsNullOrEmpty(nWirePin.FlowControlAction) ? "" : "_" + nWirePin.FlowControlAction);
            flowRow.Opcode = OpCode.Test;
            flowRow.FailAction = NwireSingleton.NwireFlag;
            if (TestPlanStatic.JobInfoSheet != null)
            {
                flowRow.Job = string.Join(",", TestPlanStatic.JobInfoSheet.Rows.Where(x => x.TesterType.Equals(testType)).Select(x => x.JobName));
            }

            flow.AddRow(flowRow);

            flowRow = new FlowRow { Opcode = OpCode.BinTable, Parameter = NwireNaming.GetBinTableName() };
            flow.AddRow(flowRow);
        }


    }
}
