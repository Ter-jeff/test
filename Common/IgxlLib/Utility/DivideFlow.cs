using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace IgxlLib.Utility
{
    public static partial class DivideFlow
    {
        [GeneratedRegex(@"(?<indexSheetName>_\d+)", RegexOptions.Compiled)]
        private static partial Regex SheetIndexSuffixRegex();
        [GeneratedRegex("if|else|endif|for|next|test", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex LoopControlOpcodeRegex();
        [GeneratedRegex("test|if|for", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex SplitPointOpcodeRegex();

        private static readonly Regex _sheetIndexSuffixRegex = SheetIndexSuffixRegex();
        private static readonly Regex _loopControlOpcodeRegex = LoopControlOpcodeRegex();
        private static readonly Regex _splitPointOpcodeRegex = SplitPointOpcodeRegex();

        public static (Dictionary<string, SubFlowSheet>, Dictionary<string, List<string>>) SplitFlowSheet(Dictionary<string, SubFlowSheet> subFlowSheets, int maxDivideRow = 50000, List<string>? executionProfileNotRunFlows = null)
        {
            Dictionary<string, SubFlowSheet> tempSubFlowSheets = [];
            Dictionary<string, List<string>> expSheets = [];

            var flowNameList = subFlowSheets.Select(subflow => subflow.Key).ToList();
            bool needDivide = true;
            foreach (KeyValuePair<string, SubFlowSheet> flow in subFlowSheets)
            {
                int cnt = 0;
                if (executionProfileNotRunFlows != null && executionProfileNotRunFlows.Exists(x => x.EqualsIgnoreCase(flow.Key)))
                {
                    needDivide = false;
                }

                if (flow.Value.Rows.Count > maxDivideRow && needDivide)
                {
                    int index = GetDivideIndex(flow, maxDivideRow);
                    if (index > 0)
                    {
                        KeyValuePair<string, SubFlowSheet> newFlow = SplitFlowAtIndex(flow.Value.Name, flow, index, ref flowNameList, ref cnt, ref tempSubFlowSheets, ref expSheets);
                        int index1 = GetDivideIndex(newFlow, maxDivideRow);
                        while (index1 > maxDivideRow)
                        {
                            newFlow = SplitFlowAtIndex(flow.Value.Name, newFlow, index1, ref flowNameList, ref cnt, ref tempSubFlowSheets, ref expSheets);
                            index1 = GetDivideIndex(newFlow, maxDivideRow);
                        }
                    }
                }
            }
            return (tempSubFlowSheets, expSheets);
        }

        public static void InsExpFlowSheet(Dictionary<string, SubFlowSheet> subFlowSheets, Dictionary<string, List<string>> expSheets)
        {
            foreach (KeyValuePair<string, SubFlowSheet> flow in subFlowSheets)
            {
                foreach (KeyValuePair<string, List<string>> expSheet in expSheets)
                {
                    int searchStart = 0;
                    int index = flow.Value.Rows.FindIndex(searchStart, row => row.Opcode.EqualsIgnoreCase("call") && row.Parameter == expSheet.Key);
                    while (index != -1)
                    {
                        foreach (string expSheetName in expSheet.Value)
                        {
                            var currRow = new FlowRow(flow.Value.Rows[index])
                            {
                                Parameter = expSheetName
                            };
                            index++;
                            flow.Value.Rows.Insert(index, currRow);
                        }
                        searchStart = index + 1;
                        index = searchStart < flow.Value.Rows.Count
                            ? flow.Value.Rows.FindIndex(searchStart, row => row.Opcode.EqualsIgnoreCase("call") && row.Parameter == expSheet.Key)
                            : -1;
                    }
                }
            }
        }

        private static KeyValuePair<string, SubFlowSheet> SplitFlowAtIndex(string flowName, KeyValuePair<string, SubFlowSheet> flow, int index, ref List<string> flowNameList, ref int cnt, ref Dictionary<string, SubFlowSheet> tempSubFlowSheets, ref Dictionary<string, List<string>> expSheets)
        {
            var newFlow = new KeyValuePair<string, SubFlowSheet>();
            if (index > 0)
            {
                cnt += 1;
                string indexSheetName = _sheetIndexSuffixRegex.Match(flow.Value.Name).Groups["indexSheetName"].ToString();
                int location = string.IsNullOrEmpty(indexSheetName) ? -1 : flow.Value.Name.IndexOf(indexSheetName, StringComparison.Ordinal);
                string newSheetName = location == -1 ? flow.Value.Name + "_" + cnt : flow.Value.Name[..location] + "_" + cnt;
                while (flowNameList.Exists(x => x.EqualsIgnoreCase(newSheetName)))
                {
                    newSheetName = newSheetName + "_" + cnt;
                }
                var newSubSheet = new SubFlowSheet(newSheetName);
                newSubSheet.Rows.AddRange(flow.Value.Rows.GetRange(index, flow.Value.Rows.Count - index));
                string key = flow.Key.Replace(flow.Value.Name, newSheetName);
                newSubSheet.SplitFromSheet = flowName;
                newFlow = new KeyValuePair<string, SubFlowSheet>(key, newSubSheet);
                tempSubFlowSheets.Add(key, newSubSheet);
                flow.Value.Rows.RemoveRange(index, flow.Value.Rows.Count - index);

                if (!expSheets.TryGetValue(flowName, out List<string>? list))
                {
                    list = [];
                    expSheets[flowName] = list;
                }
                list.Add(newSheetName);

                flow.Value.Rows.Add(GetReturnCall());
            }
            return newFlow;
        }

        private static int GetDivideIndex(KeyValuePair<string, SubFlowSheet> flow, int maxDivideRow)
        {
            var loopList = flow.Value.Rows.Where(x => x.Opcode != null).Select((v, i) => new { opcode = v.Opcode, index = i })
                .Where(x => _loopControlOpcodeRegex.IsMatch(x.opcode)).ToList();
            int ifValue = 0;
            int forValue = 0;
            int index = 0;
            foreach (var item in loopList)
            {
                if (_splitPointOpcodeRegex.IsMatch(item.opcode) & item.index > maxDivideRow &
                    ifValue == 0 & forValue == 0)
                {
                    index = item.index;
                    break;
                }
                if (item.opcode.EqualsIgnoreCase("if"))
                {
                    ifValue++;
                }
                else if (item.opcode.EqualsIgnoreCase("endif"))
                {
                    ifValue--;
                }
                else if (item.opcode.EqualsIgnoreCase("for"))
                {
                    forValue++;
                }
                else if (item.opcode.EqualsIgnoreCase("next"))
                {
                    forValue--;
                }
            }
            return index;
        }

        private static FlowRow GetReturnCall()
        {
            return new FlowRow { Opcode = "Return" };
        }
    }
}
