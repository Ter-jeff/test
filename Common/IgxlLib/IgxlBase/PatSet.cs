using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Regs;

namespace IgxlLib.IgxlBase
{
    public class PatSet : IgxlRow
    {
        public string PatSetName { get; set; } = string.Empty;

        public List<PatSetRow> PatSetRows { get; set; } = [];

        public string Domain { get; set; } = string.Empty;

        public void AddRow(PatSetRow patSetRow)
        {
            PatSetRows.Add(patSetRow);
        }

        public static string GetNewPatSetNameWithX(List<string> patterns)
        {
            if (patterns.Count == 0)
            {
                return "";
            }

            char[] arr = patterns.First().ToCharArray();
            foreach (string pattern in patterns)
            {
                for (int index = 0; index < patterns.First().Length; index++)
                {
                    if (index >= pattern.Length)
                    {
                        arr[index] = 'X';
                    }
                    else
                    {
                        if (!pattern[index].ToString().EqualsIgnoreCase(patterns.First()[index].ToString()))
                        {
                            arr[index] = 'X';
                        }
                    }
                }
            }
            return string.Join("", arr);
        }

        public static string GetPatSetName(List<string> patterns, bool onlyPl = true)
        {
            if (patterns.Count == 0)
            {
                return "";
            }

            if (patterns.Count == 1)
            {
                return patterns[0];
            }

            if (onlyPl)
            {
                var list = patterns.Where(x => !x.ContainsIgnoreCase("init") && !Reg.InitRegex.IsMatch(x.Split('_')[3])).ToList();
                if (list.Count > 0)
                {
                    patterns = list;
                }
            }
            return GetNewPatSetName(patterns);
        }

        public static string GetNewPatSetName(List<string> patterns)
        {
            int max = patterns.Max(x => x.Split('_').Length);

            bool[] arr = new bool[max];
            string[] first = patterns.First().Split('_');
            foreach (string pattern in patterns)
            {
                string[] items = pattern.Split('_');
                for (int index = 0; index < max; index++)
                {
                    if (index < items.Length && index < first.Length)
                    {
                        if (!first[index].EqualsIgnoreCase(items[index]))
                        {
                            arr[index] = true;
                        }
                    }
                    else
                    {
                        arr[index] = true;
                    }
                }
            }

            string[] final = new string[max];
            for (int index = 0; index < max; index++)
            {
                var one = new List<string>();
                foreach (string pattern in patterns)
                {
                    string[] items = pattern.Split('_');
                    if (index < items.Length)
                    {
                        if (arr[index] && !one.Any(x => x.EqualsIgnoreCase(items[index])))
                        {
                            one.Add(items[index]);
                        }

                        if (one.Count == 0)
                        {
                            one.Add(items[index]);
                        }
                    }
                    final[index] = string.Join("_", one);
                }
            }
            return string.Join("_", final);
        }
    }
}
