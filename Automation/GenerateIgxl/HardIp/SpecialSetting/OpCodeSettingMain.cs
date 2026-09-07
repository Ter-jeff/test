using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.HardIp.SpecialSetting
{
    public static class OpCodeSettingMain
    {
        public static List<FlowRow> GenOpcodeSetting(List<string> opSettings, string blockName = "", string voltage = "", string enable = "")
        {
            var reg = new Regex(@"\w+", RegexOptions.IgnoreCase);
            string voltageFlag = GenVoltageFlag(voltage);

            var flowRows = new List<FlowRow>();
            foreach (string setting in opSettings)
            {
                var opcodeRow = new FlowRow { Opcode = setting.Split(':')[0] };
                if (!Regex.IsMatch(opcodeRow.Opcode, "elseif|endif|if", RegexOptions.IgnoreCase))
                {
                    opcodeRow.Enable = enable;
                }

                opcodeRow.Parameter = setting.Split(':')[1];
                opcodeRow.Parameter = reg.Replace(opcodeRow.Parameter, delegate (Match m)
                {
                    if (Regex.IsMatch(m.Value, "^pp_|^dd_", RegexOptions.IgnoreCase))
                    {
                        return "F_" + blockName + "_" + m.Value + voltageFlag;
                    }

                    return m.Value;
                });

                flowRows.Add(opcodeRow);
            }
            return flowRows;
        }

        private static string GenVoltageFlag(string labelVoltage)
        {
            const string flagN = "_N_Flag";
            const string flagL = "_L_Flag";
            const string flagH = "_H_Flag";
            if (string.IsNullOrEmpty(labelVoltage))
            {
                return string.Empty;
            }

            switch (labelVoltage)
            {
                case HardIpConstData.LabelNv:
                    return flagN;
                case HardIpConstData.LabelLv:
                    return flagL;
                case HardIpConstData.LabelHv:
                    return flagH;
                default:
                    return string.Empty;
            }
        }
    }
}
