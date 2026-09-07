using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.SpecialSetting
{
    public class RelaySettingMain
    {
        public static List<FlowRow> GenRelaySetting(SubFlowSheet flowSheet, Dictionary<string, string> relaySetting, string enable = "", bool needreverse = false)
        {
            var flowRows = new List<FlowRow>();
            foreach (KeyValuePair<string, string> item in relaySetting)
            {
                string job = item.Key;
                string settingInJob = needreverse ? SearchInfo.ReverseRelaySetting(item.Value) : item.Value;
                flowRows.AddRange(GenRelaySettingInJob(flowSheet, settingInJob, job, enable));
            }
            return flowRows;
        }

        public static List<FlowRow> GenRelaySettingInJob(SubFlowSheet flowSheet, string relaySetting, string job, string enable = "")
        {
            var flowRows = new List<FlowRow>();
            foreach (string setting in relaySetting.Split(';'))
            {
                var relayRow = new FlowRow
                {
                    Enable = enable,
                    Opcode = OpCode.Test,
                    Parameter = HardIpConstData.PrefixAtgRelay + setting.Replace(":", "_").Replace("&", "_"),
                    Job = job
                };

                flowRows.Add(relayRow);
                flowSheet?.AddRow(relayRow);
            }
            return flowRows;
        }
    }
}
