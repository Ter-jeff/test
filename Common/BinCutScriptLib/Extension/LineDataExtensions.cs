using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

using CommonLib.Extension;

namespace BinCutScriptLib.Extension
{
    public static class LineDataExtensions
    {
        public static List<List<BinCutLineBase>> GetDigSrcLines(this List<BinCutLineBase> binCutLineBases, string split)
        {
            List<List<BinCutLineBase>> result = [];
            List<BinCutLineBase> lineDatas = [];
            for (int i = 0; i < binCutLineBases.Count; i++)
            {
                BinCutLineBase line = binCutLineBases[i];
                if (i != 0 && line.Line.StartsWithIgnoreCase(split))
                {
                    result.Add(lineDatas);
                    lineDatas = [];
                }
                lineDatas.Add(line);
            }
            result.Add(lineDatas);
            return result;
        }

        public static List<List<BinCutLineBase>> GetInstanceines(this List<BinCutLineBase> binCutLineBases)
        {
            // Number Site  Test Name Pattern
            //<MS004_SocTd_MSX004_SOC_MSX004_TD_CCA0_PL00_SSC_BV_Csharp>
            // 42000000     0     MS004_SocTd_MSX004_SOC_MSX004_TD_CCA0_PL00_SSC_BV_Csharp 
            // 42000000     1     MS004_SocTd_MSX004_SOC_MSX004_TD_CCA0_PL00_SSC_BV_Csharp 
            // 42000000     2     MS004_SocTd_MSX004_SOC_MSX004_TD_CCA0_PL00_SSC_BV_Csharp 
            // 42000000     3     MS004_SocTd_MSX004_SOC_MSX004_TD_CCA0_PL00_SSC_BV_Csharp 

            // Number Site  Test Name Pin Channel  Low Measured             High Force          Loc
            // 42000055     0     EQN VDD_SOC                        9.e246   0              8                    11             0              0
            // 42000055     1     EQN VDD_SOC                        9.e446   0              8                    11             0              0
            List<List<BinCutLineBase>> result = [];
            List<BinCutLineBase> lineDatas = [];
            bool flag = false;
            for (int i = 0; i < binCutLineBases.Count; i++)
            {
                BinCutLineBase line = binCutLineBases[i];
                if (flag && !Reg.RegexLine.IsMatch(line.Line.Trim()) && !line.Line.StartsWith('<'))
                {
                    result.Add(lineDatas);
                    lineDatas = [];
                    flag = false;
                }
                if (Reg.RegexLine1.IsMatch(line.Line))
                {
                    flag = true;
                }
                lineDatas.Add(line);
            }
            result.Add(lineDatas);
            return result;
        }

        public static FlowNode? GetFlowTree(this List<BinCutLineBase> binCutLineBases)
        {
            if (binCutLineBases == null || binCutLineBases.Count == 0)
            {
                return null;
            }

            Stack<FlowNode> stack = new Stack<FlowNode>();
            FlowNode? root = null;
            for (int index = 0; index < binCutLineBases.Count; index++)
            {
                BinCutLineBase line = binCutLineBases[index];
                string text = line.Line.Trim();
                if ((text.StartsWith("*print: ") && text.EndsWith("start*")) || (text.StartsWith("Flow ") && text.EndsWith("Start ")))
                {
                    int startIdx = 8;
                    int endIdx = text.Length - 7;
                    string flowName = text[startIdx..endIdx].Trim();
                    var node = new FlowNode { Line = line, FlowName = flowName };

                    if (stack.Count == 0)
                    {
                        root = node;
                    }
                    else
                    {
                        stack.Peek().Children.Add(node);
                    }

                    stack.Push(node);
                }
                else if ((text.StartsWith("*print: ") && text.EndsWith("end*")) || (text.StartsWith("Flow ") && text.EndsWith("Stop ")))
                {
                    if (stack.Count > 0)
                    {
                        stack.Pop();
                    }
                }
                else
                {
                    if (stack.Count > 0)
                    {
                        if (!line.Line.All(x => x == '*'))
                        {
                            stack.Peek().Children.Add(new Node { Line = line });
                        }
                    }
                }
            }

            return root;
        }

        public static bool GetAndRemoveRange(this List<BinCutLineBase> lines, out List<BinCutLineBase> result, string startText, string endText)
        {
            result = [];
            if (lines.Count == 0)
            {
                return false;
            }

            int start = lines.FindIndex(x => x.Line.Contains(startText));
            if (start == -1)
            {
                return false;
            }

            int end = lines.FindIndex(start + 1, x => x.Line.Contains(endText));
            if (end == -1)
            {
                return false;
            }

            result = lines.GetRange(start, end - start + 1);
            lines.RemoveRange(0, end + 1);
            return true;
        }
    }
}
