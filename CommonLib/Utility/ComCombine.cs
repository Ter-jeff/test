using System;
using System.Collections.Generic;
using System.Linq;

namespace CommonLib.Utility
{
    public class ComCombine
    {
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
            if (!list.Any())
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

            if (str1.Equals(str2, StringComparison.CurrentCultureIgnoreCase))
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

        public static string CombineEnableWordByOr(string str1, string str2)
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

            if (str1.Equals(str2, StringComparison.CurrentCultureIgnoreCase))
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

            return newStr1 + " || " + newStr2;
        }
    }
}
