using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.PostAction.TempMon.Data;
using Automation.GenerateIgxl.PostAction.TempMon.Enums;

namespace Automation.Utility.Basic
{
    public static class TempMonService
    {
        private static Regex _regexTempMonSyntax = new Regex(@"TempMon_(?<mode>[^:]+):(?<condition>Enable|Disable)(?=\s|,|;|$)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool TrySetTempMon(HashSet<TempMonData> tempMonDatas, string tempMonSyntax, string item, EnumType type)
        {
            if (TryGetTempMonSyntax(tempMonSyntax, out string mode, out EnumCondition condition))
            {
                SetTempMon(tempMonDatas, mode, condition, type, item);
                return true;
            }
            return false;
        }

        public static void SetTempMon(HashSet<TempMonData> tempMonDatas, string mode, EnumCondition condition, EnumType type, string item)
        {
            TempMonData data = new TempMonData() { Item = item, Condition = condition, Type = type, Mode = mode };
            tempMonDatas.Add(data);
        }

        public static bool TryGetTempMonSyntax(string tempMonSyntax, out string mode, out EnumCondition condition)
        {
            mode = "";
            condition = EnumCondition.Unknown;

            if (string.IsNullOrEmpty(tempMonSyntax))
            {
                return false;
            }

            Match tempMonMatch = _regexTempMonSyntax.Match(tempMonSyntax);
            if (tempMonMatch.Success)
            {
                mode = tempMonMatch.Groups["mode"].Value.ToUpper();
                string enable = tempMonMatch.Groups["condition"].Value;
                if (enable.Equals("Disable", StringComparison.OrdinalIgnoreCase))
                {
                    condition = EnumCondition.Exclude;
                }
                else if (enable.Equals("Enable", StringComparison.OrdinalIgnoreCase))
                {
                    condition = EnumCondition.Include;
                }
                return true;
            }
            return false;
        }
    }
}
