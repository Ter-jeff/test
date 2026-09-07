using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Datalog;
using CommonLib.Extension;

using EfuseCheckCmdLib.Datalog;
using EfuseCheckCmdLib.DataStructure;

using TestPlanLib.Efuse;

namespace EfuseCheckCmdLib.Utility
{
    public static partial class LineHandle
    {
        private static readonly Regex _regNumber = NumberRegex();
        private static readonly Regex _regCfgScenario = CfgScenarioRegex();

        [GeneratedRegex(@"^\d+$")]
        private static partial Regex NumberRegex();

        [GeneratedRegex(@"CFG_[A-Z]\d+", RegexOptions.IgnoreCase)]
        private static partial Regex CfgScenarioRegex();

        public static void HandleSetWriteVariableLine(List<SetWriteVariableLine> setWriteVariableLines, string line)
        {
            var data = new SetWriteVariableLine(line);
            SetWriteVariableLine? target = setWriteVariableLines.Find(x => x.Site == data.Site && x.Key.EqualsIgnoreCase(data.Key));
            if (target != null)
            {
                target.Value = data.Value;
            }
            else
            {
                setWriteVariableLines.Add(data);
            }
        }

        public static void HandleEfuseLine(List<EfuseRow> efuseRows, int i, string line, bool hasReadWaferData = false)
        {
            EfuseRow? efuseRow = new EfuseLine(line, i).ToRow(hasReadWaferData);
            if (efuseRow != null)
            {
                efuseRows.Add(efuseRow);
            }
        }

        public static void HandleLimitStart(List<LimitRow> limitRows, ref bool hasLimitStart, ref int startIndex, List<string> lines, int i, string[] arr)
        {
            if (!_regNumber.IsMatch(arr[0]) && !arr[0].StartsWith("[INFO]"))
            {
                hasLimitStart = false;
                List<string> all = lines.GetRange(startIndex + 1, i - startIndex - 1);
                List<LimitLine> limitLines = [];
                for (int j = 0; j < all.Count; j++)
                {
                    string item = all[j];
                    limitLines.Add(new LimitLine { Line = item, LineNo = startIndex + 1 + 1 + j });
                }
                limitRows.AddRange(limitLines.Select(x => x.ToRow()).Where(x => x != null)!);
                startIndex = -1;
            }
        }

        public static void HandleEnableWords(EfuseDramTable efuseDramTable, ref string scenario, ref string dramType, string line)
        {
            string[] array = line.Split([' ', '\t', ':', '|', '\'']);
            foreach (string item in array)
            {
                if (_regCfgScenario.IsMatch(item))
                {
                    scenario = item.Replace("CFG_", "");
                    break;
                }
            }
            foreach (string item in array)
            {
                foreach (string dRAMType in efuseDramTable.Titles)
                {
                    if (item.EqualsIgnoreCase(dRAMType))
                    {
                        dramType = item;
                        break;
                    }
                }
            }
        }
    }
}
