using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CommonLib.Extension
{
    public static partial class StringExtensions
    {
        [GeneratedRegex(@"\d", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexHasDigit();
        [GeneratedRegex(@"(?<cores>\[.+\])", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexBracketGroup();
        [GeneratedRegex(@"(?<cores>\d+:\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexColonRange();
        [GeneratedRegex(@"(?<cores>\d+..\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexDotRange();

        private static readonly Regex _regexHasDigit = MyRegexHasDigit();
        private static readonly Regex _regexBracketGroup = MyRegexBracketGroup();
        private static readonly Regex _regexColonRange = MyRegexColonRange();
        private static readonly Regex _regexDotRange = MyRegexDotRange();

        public static GroupItem GetIterates(this string text)
        {
            var groupItem = new GroupItem { OriText = text };
            if (!_regexBracketGroup.IsMatch(text))
            {
                return groupItem;
            }

            var matches = new List<string>();
            matches = GetMatches(ref matches, text, '[', ']');
            var items = new Dictionary<int, List<Tuple<string, string>>>();
            foreach (string match in matches)
            {
                string[] codeArray = match.TrimStart('[').TrimEnd(']').Split(',');
                for (int j = 0; j < codeArray.Length; j++)
                {
                    string item = codeArray[j];
                    if (_regexColonRange.IsMatch(item) || _regexDotRange.IsMatch(item))
                    {
                        string[] arr = item.Split([":", ".."], StringSplitOptions.None);
                        if (int.TryParse(arr.First(), out int start) &&
                            int.TryParse(arr.Last(), out int end))
                        {
                            if (start > end)
                            {
                                for (int i = start; i >= end; i--)
                                {
                                    int key = i - start;
                                    if (items.TryGetValue(key, out List<Tuple<string, string>>? value))
                                    {
                                        value.Add(new Tuple<string, string>(match, i.ToString()));
                                    }
                                    else
                                    {
                                        items.Add(key, [new(match, i.ToString())]);
                                    }
                                }
                            }
                            else
                            {
                                for (int i = start; i <= end; i++)
                                {
                                    int key = i - start;
                                    if (items.TryGetValue(key, out List<Tuple<string, string>>? value))
                                    {
                                        value.Add(new Tuple<string, string>(match, i.ToString()));
                                    }
                                    else
                                    {
                                        items.Add(key, [new(match, i.ToString())]);
                                    }
                                }
                            }
                        }
                    }
                    else if (_regexHasDigit.IsMatch(item))
                    {
                        if (items.TryGetValue(0, out List<Tuple<string, string>>? value))
                        {
                            value.Add(new Tuple<string, string>(match, item));
                        }
                        else
                        {
                            items.Add(0, [new(match, item)]);
                        }
                    }
                }
            }
            groupItem.Groups = items;
            return groupItem;
        }

        private static List<string> GetMatches(ref List<string> items, string text, char first, char last)
        {
            for (int i = 0; i < text.Length; i++)
            {
                string curText = text[..(i + 1)];
                int firstCnt = curText.Count(x => x == first);
                int lastCnt = curText.Count(x => x == last);
                if (firstCnt != 0 && firstCnt == lastCnt)
                {
                    int start = curText.IndexOf(first);
                    items.Add(curText.Substring(start, i - start + 1));
                    GetMatches(ref items, text.Substring(i + 1, text.Length - i - 1), first, last);
                    break;
                }
            }
            return items;
        }
    }

    public class GroupItem
    {
        public string OriText = "";
        public Dictionary<int, List<Tuple<string, string>>> Groups = [];
    }
}
