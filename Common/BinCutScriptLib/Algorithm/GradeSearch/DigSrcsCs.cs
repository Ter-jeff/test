using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Base.Line;
using BinCutScriptLib.Static;

namespace BinCutScriptLib.Algorithm.GradeSearch
{
    internal class DigSrcsCs
    {
        private enum EnumFoundSelsramMethod
        {
            LogLevelDebug,
            LogRegexSelsramAssignment,
            None
        }

        private readonly List<DigSrcCs> _digSrcss = [];
        private readonly List<BinCutLineBase> _lines;

        public DigSrcsCs(List<BinCutLineBase> binCutLineBases)
        {
            _lines = binCutLineBases;
            ReadLines();
        }

        private void ReadLines()
        {
            var digSrcCs = new DigSrcCs();
            bool flag = false;
            foreach (BinCutLineBase line in _lines)
            {
                if (line.Line.Contains("Pattern ="))
                {
                    if (flag)
                    {
                        _digSrcss.Add(digSrcCs);
                        digSrcCs = new DigSrcCs();
                    }
                    digSrcCs.Pattern = line.Line.Split('=')[1].Trim();
                    flag = true;
                }
                else if (line.Line.Contains("Src Bits ="))
                {
                    digSrcCs.SrcBits = int.Parse(line.Line.Split('=')[1].Trim());
                }
                else if (line.Line.Contains("SrcPin ="))
                {
                    digSrcCs.SrcPin = line.Line.Split('=')[1].Trim();
                }
                else if (line.Line.Contains("DataSequence:"))
                {
                    digSrcCs.DataSequence = line.Line.Split(':')[1].Trim();
                }
                else if (line.Line.Contains("Assignment:"))
                {
                    digSrcCs.Assignment = line.Line.Split(':')[1].Trim();
                }
                else if (line.Line.Contains("[Site"))
                {
                    digSrcCs.OutputStrings.Add(line);
                }
                else if (line.Line.Contains("Output String"))
                {
                    digSrcCs.OutputString = line.Line;
                }
            }
            _digSrcss.Add(digSrcCs);
        }

        internal List<HarvestSourceCodetRow> GetDsscsCsForHarv()
        {
            var rows = new List<HarvestSourceCodetRow>();
            //var found = DigSrcss.Find(x => x.Assignment.StartsWith("B=Dssc", StringComparison.CurrentCultureIgnoreCase));
            DigSrcCs? found = _digSrcss.Find(x => x?.Assignment != null && Reg.RegexHarvestSourceCodeAssignment.IsMatch(x.Assignment));

            if (found != null)
            {
                foreach (BinCutLineBase outputString in found.OutputStrings)
                {
                    string site = Reg.RegexSelSram.Match(outputString.Line).Groups["site"].ToString();
                    string selSramBitStr = Reg.RegexSelSram.Match(outputString.Line).Groups["bitString"].ToString();
                    var row = new HarvestSourceCodetRow
                    {
                        SrcBit = found.SrcBits.ToString()
                    };
                    _ = int.TryParse(site, out row.Site);
                    row.MainFlag = "";
                    row.HarvestSourceCode = selSramBitStr;
                    List<string> patList = [.. found.Pattern.Split('_')];
                    //Remove Pattern version
                    if (patList.Count > 4)
                    {
                        patList.RemoveRange(patList.Count - 3, 3);
                    }

                    row.PatternName = string.Join("_", patList);
                    //Site:0,F_GFX_CORE0=False, Src Bits = 34, HarvestSourceCode [ First(L) ==> Last(R) ] :0010010011000011001001001100100100
                    string order = "First(L) ==> Last(R)";
                    if (found.OutputString != "[INFO]  Output String [ LSB(L) ==> MSB(R) ]:")
                    {
                        order = "First(R) ==> Last(L)";
                    }

                    string text = $"Site:{site}, Src Bits = {found.SrcBits}, HarvestSourceCode {order} :{selSramBitStr}";
                    row.Line = new HarvestSourceCodetLine() { Line = text, LineNo = outputString.LineNo };
                    Match match = Reg.RegexAssignment.Match(found.Assignment);
                    if (match.Success)
                    {
                        string result = match.Groups[1].Value;
                        row.Assignment = result;
                    }
                    row.Label = "";
                    rows.Add(row);
                }
            }
            return rows;
        }

        internal List<HarvestSourceCodetRow> GetDsscsCs()
        {
            var rows = new List<HarvestSourceCodetRow>();
            string text = "";
            for (int i = 0; i < _digSrcss.Count; i++)
            {
                List<BinCutLineBase> outputStrings = _digSrcss[i].OutputStrings;
                for (int j = 0; j < outputStrings.Count; j++)
                {
                    string outputString = outputStrings[j].Line;
                    if (outputString.Contains("DSSC"))
                    {
                        string[] arr = [.. outputString.Split([' ', '(', ')', ']'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim())];
                        string site = arr[2];
                        var row = new HarvestSourceCodetRow
                        {
                            SrcBit = "24",
                            Label = arr[5],
                            MainFlag = "",
                            HarvestSourceCode = arr[3]
                        };
                        text = $"Site:{site}, Src Bits = {row.SrcBit}, HarvestSourceCode DefaultOrder :{row.HarvestSourceCode}";
                        row.Line = new HarvestSourceCodetLine() { Line = text, LineNo = i };
                        rows.Add(row);
                    }
                }
            }
            return rows;
        }

        internal List<BinCutLineBase> GetSelsramsCs()
        {
            //[INFO]  [Site 0] 1111(C)
            //[INFO]  [Site 2] 0101(Selsram())
            //[INFO]  [Site 0] 1(disp_sram_rail),1(gpu_sram_rail),1(dcs_sram_rail),1(soc_sram_rail)

            var rows = new List<BinCutLineBase>();
            EnumFoundSelsramMethod method = EnumFoundSelsramMethod.None;
            DigSrcCs? found = _digSrcss.Find(x => x.Assignment.Contains("_sram_"));
            if (found != null)
            {
                method = EnumFoundSelsramMethod.LogLevelDebug;
            }
            else
            {
                found = _digSrcss.Find(x => Reg.RegexSelsramAssignment.IsMatch(x.Assignment));
                if (found != null)
                {
                    method = EnumFoundSelsramMethod.LogRegexSelsramAssignment;
                }
            }
            string siteString = "";

            //var found = DigSrcss.Find(x => Reg.RegexSelsramAssignment.IsMatch(x.Assignment));
            if (found != null)
            {
                for (int i = 0; i < found.OutputStrings.Count; i++)
                {
                    switch (method)
                    {
                        case EnumFoundSelsramMethod.LogLevelDebug:
                            // Processing when found with "_sram_"

                            if (found.OutputStrings[i].Line.Contains("[Site "))
                            {
                                int siteStartIndex = found.OutputStrings[i].Line.IndexOf("[Site ") + 6;
                                int siteEndIndex = found.OutputStrings[i].Line.IndexOf(']', siteStartIndex);
                                siteString = found.OutputStrings[i].Line[siteStartIndex..siteEndIndex];
                            }

                            List<string> sramBits = [];
                            List<string> assignmentStrings = [];
                            string lineContent = found.OutputStrings[i].Line;

                            int infoEndIndex = lineContent.IndexOf(']') + 1;
                            if (infoEndIndex > 0)
                            {
                                lineContent = lineContent[infoEndIndex..].Trim().Replace("[Site " + siteString + "] ", "");
                            }

                            string[] segments = lineContent.Split(',');

                            foreach (string segment in segments)
                            {
                                if (segment.Contains("_sram_"))
                                {
                                    int bitEndIndex = segment.IndexOf('(');
                                    string[] arr = segment.Split('(', ')');
                                    if (bitEndIndex > 0)
                                    {
                                        string bit = segment[..bitEndIndex].Trim();
                                        sramBits.Add(bit);
                                        assignmentStrings.Add(arr[1]);
                                    }
                                }
                            }

                            var row = new BinCutLineBase
                            {
                                Line = $"[INFO]  [Site {siteString}] {lineContent}",
                                LineNo = found.OutputStrings[i].LineNo
                            };
                            rows.Add(row);
                            break;

                        case EnumFoundSelsramMethod.LogRegexSelsramAssignment:

                            string site = Reg.RegexSelSram.Match(found.OutputStrings[i].Line).Groups["site"].ToString();
                            lineContent = found.OutputStrings[i].Line.Split(['[', ']', ' '], StringSplitOptions.RemoveEmptyEntries).Last();
                            var row2 = new BinCutLineBase
                            {
                                Line = $"[INFO]  [Site {site}] {lineContent}",
                                LineNo = found.OutputStrings[i].LineNo
                            };
                            rows.Add(row2);
                            break;

                        case EnumFoundSelsramMethod.None:
                            break;
                    }
                }
            }
            return rows;
        }
    }
}
