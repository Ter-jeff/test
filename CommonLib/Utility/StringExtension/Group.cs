using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CommonLib.Utility.StringExtension
{
    public static class Group
    {
        private static readonly Regex _regex = new Regex(@"\d", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(@"(?<cores>\[.+\])", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex3 = new Regex(@"(?<cores>\d+:\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex4 = new Regex(@"(?<cores>\d+..\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static GroupItem GetIterates(this string text)
        {
            var groupItem = new GroupItem();
            groupItem.OriText = text;
            if (!_regex2.IsMatch(text))
            {
                return groupItem;
            }

            var matches = new List<string>();
            matches = GetMatches(ref matches, text, '[', ']');
            var items = new Dictionary<int, List<Tuple<string, string>>>();
            foreach (string match in matches)
            {
                string[] codeArray = match.TrimStart('[').TrimEnd(']').Split(',');
                for (int j = 0; j < codeArray.Count(); j++)
                {
                    string item = codeArray[j];
                    if (_regex3.IsMatch(item) || _regex4.IsMatch(item))
                    {
                        int start;
                        int end;
                        string[] arr = item.Split(new[] { ":", ".." }, StringSplitOptions.None);
                        if (int.TryParse(arr.First(), out start) &&
                            int.TryParse(arr.Last(), out end))
                        {
                            if (start > end)
                            {
                                for (int i = start; i >= end; i--)
                                {
                                    int key = i - start;
                                    if (items.ContainsKey(key))
                                    {
                                        items[key].Add(new Tuple<string, string>(match, i.ToString()));
                                    }
                                    else
                                    {
                                        items.Add(key, new List<Tuple<string, string>> { new Tuple<string, string>(match, i.ToString()) });
                                    }
                                }
                            }
                            else
                            {
                                for (int i = start; i <= end; i++)
                                {
                                    int key = i - start;
                                    if (items.ContainsKey(key))
                                    {
                                        items[key].Add(new Tuple<string, string>(match, i.ToString()));
                                    }
                                    else
                                    {
                                        items.Add(key, new List<Tuple<string, string>> { new Tuple<string, string>(match, i.ToString()) });
                                    }
                                }
                            }
                        }
                    }
                    else if (_regex.IsMatch(item))
                    {
                        if (items.ContainsKey(0))
                        {
                            items[0].Add(new Tuple<string, string>(match, item));
                        }
                        else
                        {
                            items.Add(0, new List<Tuple<string, string>> { new Tuple<string, string>(match, item) });
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
                string curText = text.Substring(0, i + 1);
                int fisrtCnt = curText.Count(x => x == first);
                int lasttCnt = curText.Count(x => x == last);
                if (fisrtCnt != 0 && fisrtCnt == lasttCnt)
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
        public string OriText;
        public Dictionary<int, List<Tuple<string, string>>> Groups = new Dictionary<int, List<Tuple<string, string>>>();
    }
}
