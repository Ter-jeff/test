using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

namespace CommonLib.Utility
{
    public static class Combination
    {
        public static List<List<string>> GetAllPossibleCombos(List<List<string>> items)
        {
            IEnumerable<List<string>> combos = [[]];
            foreach (List<string> item in items)
            {
                combos = [.. combos.SelectMany(r => item.Select(x =>
                {
                    var n = new List<string>(r);

                    if (x != null)
                    {
                        n.Add(x);
                    }

                    return n;
                }))];
            }

            return [.. combos.Where(x => x.Count > 0)];
        }

        public static List<List<List<string>>> GetAllPossibleCombos(List<List<List<string>>> items)
        {
            IEnumerable<List<List<string>>> combos = [[]];

            foreach (List<List<string>> item in items)
            {
                combos = [.. combos.SelectMany(r => item.Select(x =>
                {
                    List<List<string>> n = [.. r.Select(innerList => new List<string>(innerList))];

                    if (x != null)
                    {
                        n.Add([.. x]);
                    }
                    return n;
                }))];
            }

            var result = combos.Where(x => x.Count > 0).ToList();

            if (result.Count == 0)
            {
                return [[]];
            }

            return result;
        }

        public static string CombineByUnderLine(string str1, string str2)
        {
            str1 = str1.Trim();
            str2 = str2.Trim();
            if (string.IsNullOrEmpty(str1))
            {
                return str2;
            }

            if (string.IsNullOrEmpty(str2))
            {
                return str1;
            }

            return str1 + "_" + str2;
        }

        public static string CombineByUnderLine(List<string> list)
        {
            if (list.Count == 0)
            {
                return "";
            }

            return string.Join("_", list.Where(x => !string.IsNullOrEmpty(x)).ToList());
        }

        public static string CombineEnableWord(string str1, string str2)
        {
            str1 = str1.Trim();
            str2 = str2.Trim();
            if (string.IsNullOrEmpty(str1))
            {
                return str2;
            }

            if (string.IsNullOrEmpty(str2))
            {
                return str1;
            }

            if (str1.EqualsIgnoreCase(str2))
            {
                return str1;
            }

            string newStr1 = str1;
            if (str1.Contains('|') || str1.Contains('&'))
            {
                newStr1 = "(" + str1 + ")";
            }

            string newStr2 = str2;
            if (str2.Contains('|') || str2.Contains('&'))
            {
                newStr2 = "(" + str2 + ")";
            }

            return newStr1 + " && " + newStr2;
        }
    }
}
