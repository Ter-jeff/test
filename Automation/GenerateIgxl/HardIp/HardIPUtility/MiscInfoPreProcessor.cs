using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;

namespace Automation.GenerateIgxl.HardIp.HardIPUtility
{
    public static class MiscInfoPreProcessor
    {
        public static void PreProcessMiscInfo(Dictionary<string, HardIpSheet> planDic)
        {
            foreach (KeyValuePair<string, HardIpSheet> sheetItem in planDic)
            {
                GetTempOpcode(sheetItem, HardIpConstData.HighTemp, "Opcode:if:Temp_85C");
                GetTempOpcode(sheetItem, HardIpConstData.RoomTemp, "Opcode:if:Temp_25C");
                GetCondPatternOpcode(sheetItem, HardIpConstData.ExecCond);
            }
        }

        private static void GetTempOpcode(KeyValuePair<string, HardIpSheet> sheetItem, string condition, string opcode)
        {
            bool last = false;
            for (int i = 0; i < sheetItem.Value.Rows.Count; i++)
            {
                bool current = sheetItem.Value.Rows[i].MiscInfoDict.ContainsKey(condition);
                if (current)
                {
                    if (i + 1 < sheetItem.Value.Rows.Count)
                    {
                        bool next = sheetItem.Value.Rows[i + 1].MiscInfoDict.ContainsKey(condition);
                        if (!last)
                        {
                            sheetItem.Value.Rows[i].MiscInfo += ";" + opcode;
                        }

                        if (!next)
                        {
                            sheetItem.Value.Rows[i].MiscInfo += ";" + "Opcode:EndIf";
                        }
                    }
                    else
                    {
                        if (!last)
                        {
                            sheetItem.Value.Rows[i].MiscInfo += ";" + opcode;
                        }

                        sheetItem.Value.Rows[i].MiscInfo += ";" + "Opcode:EndIf";
                    }
                }
                last = current;
            }
        }

        private static void GetCondPatternOpcode(KeyValuePair<string, HardIpSheet> sheetItem, string condition)
        {
            bool last = false;
            string lastExecCond = "";
            for (int i = 0; i < sheetItem.Value.Rows.Count; i++)
            {
                bool current = sheetItem.Value.Rows[i].MiscInfoDict.ContainsKey(condition);
                string execCond = "";
                if (current)
                {
                    execCond = GetExecCond(sheetItem.Value.Rows[i].MiscInfoDict, condition);
                    if (execCond != "" && execCond != lastExecCond)
                    {
                        if (i + 1 < sheetItem.Value.Rows.Count)
                        {
                            bool next = sheetItem.Value.Rows[i + 1].MiscInfoDict.ContainsKey(condition);
                            string nextExecCond = GetExecCond(sheetItem.Value.Rows[i + 1].MiscInfoDict, condition);
                            if (!last)
                            {
                                sheetItem.Value.Rows[i].MiscInfo += ";" + "Opcode:if:" + execCond;
                            }

                            if (!next || execCond != nextExecCond)
                            {
                                sheetItem.Value.Rows[i].MiscInfo += ";" + "Opcode:EndIf";
                            }
                        }
                        else
                        {
                            if (!last)
                            {
                                sheetItem.Value.Rows[i].MiscInfo += ";" + "Opcode:if:" + execCond;
                            }

                            sheetItem.Value.Rows[i].MiscInfo += ";" + "Opcode:EndIf";
                        }
                    }
                }
                last = current;
                lastExecCond = execCond;
            }
        }

        private static string GetExecCond(Dictionary<string, string> miscInfoDict, string condition)
        {
            string execCond = "";
            if (miscInfoDict.TryGetValue(condition, out string value))
            {
                string execItem = value;  //one cell only can use one execcond
                execCond = !execItem.Equals(condition, System.StringComparison.OrdinalIgnoreCase) ? execItem : "";
            }
            return execCond;
        }
    }
}
