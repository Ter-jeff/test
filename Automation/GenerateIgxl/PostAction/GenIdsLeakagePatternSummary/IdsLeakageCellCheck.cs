using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;

namespace CommonLib.IdsLeakageCell
{
    [ExcludeFromCodeCoverage]
    public class IdsLeakageCellCheck
    {
        public List<List<string>> WorkFlow(PatternResult patternResult, string xpinStr, string unUsedPinStr, string comment = null)
        {
            var xpinList = new List<string>();
            var unUsedList = new List<string>();
            var extraList = new List<string>();
            var data = new List<List<string>>();


            if (xpinStr.Length >= 8000 || unUsedPinStr.Length >= 8000 || (comment != null && comment.Length >= 8000))
            {
                if (xpinStr.Length >= 8000)
                {
                    ChunkPins(xpinStr, xpinList);
                }
                else
                {
                    xpinList.Add(xpinStr);
                }


                if (unUsedPinStr.Length >= 8000)
                {
                    ChunkPins(unUsedPinStr, unUsedList);
                }
                else
                {
                    unUsedList.Add(unUsedPinStr);
                }

                if (comment != null && comment.Length >= 8000)
                {
                    BuildExtraList(patternResult, extraList);
                }
                else
                {
                    extraList.Add(comment);
                }


                int count = xpinList.Count > unUsedList.Count ? xpinList.Count : unUsedList.Count;
                count = count > extraList.Count ? count : extraList.Count;
                for (int i = 0; i < count; i++)
                {
                    data.Add(BuildDataRow(patternResult, xpinList, unUsedList, extraList, i));
                }
            }
            else
            {
                data.Add(new List<string> { patternResult.Pattern, patternResult.PatternVersion, xpinStr, unUsedPinStr, comment });
            }

            return data;
        }

        private void ChunkPins(string pinStr, List<string> pinList)
        {
            string[] pins = pinStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            var current = new StringBuilder();

            foreach (string pin in pins)
            {
                string token = pin.Trim();

                int additionalLength = current.Length == 0 ? token.Length : token.Length + 1;

                if (current.Length + additionalLength > 8000)
                {
                    pinList.Add(current.ToString());
                    current.Clear();
                }
                if (current.Length > 0)
                {
                    current.Append(',');
                }
                current.Append(token);
            }

            if (current.Length > 0)
            {
                pinList.Add(current.ToString());
            }
        }

        private void BuildExtraList(PatternResult patternResult, List<string> extraList)
        {
            string str = "";
            if (patternResult.MissSubFunction.Any())
            {
                str = "MissSubFunction:" + patternResult.MissSubFunction.First();
                for (int i = 1; i < patternResult.MissSubFunction.Count; i++)
                {
                    string temp = str;
                    str += "," + patternResult.MissSubFunction[i];
                    if (str.Length >= 7000)
                    {
                        extraList.Add(temp);
                        str = patternResult.MissSubFunction[i];
                    }
                }
            }

            if (patternResult.PatExtraPins.Any())
            {
                if (!string.IsNullOrEmpty(str) && patternResult.MissSubFunction.Any())
                {
                    str += ";";
                }

                str += "PatternExtraPins:" + patternResult.PatExtraPins.First();
                for (int i = 1; i < patternResult.PatExtraPins.Count; i++)
                {
                    string temp = str;
                    str += "," + patternResult.PatExtraPins[i];
                    if (str.Length >= 8000)
                    {
                        extraList.Add(temp);
                        str = patternResult.PatExtraPins[i];
                    }
                }
            }
            if (!string.IsNullOrEmpty(str))
            {
                extraList.Add(str);
            }
        }

        private List<string> BuildDataRow(PatternResult patternResult, List<string> xpinList, List<string> unUsedList, List<string> extraList, int i)
        {
            return new List<string>
            {
                i == 0 ? patternResult.Pattern : "",
                i == 0 ? patternResult.PatternVersion : "",
                xpinList.Count > 0 && xpinList.Count > i ? xpinList[i] : "",
                unUsedList.Count > 0 && unUsedList.Count > i ? unUsedList[i] : "",
                extraList.Count > 0 && extraList.Count > i ? extraList[i] : ""
            };
        }
    }
}
