using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace CommonLib.Extension
{
    public static class StringExtensions
    {
        public static string Join(this string connector, params string[] components)
        {
            return string.Join(connector, components);
        }

        public static string BitToInt(this string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            if (double.TryParse(text, out _))
            {
                return text;
            }

            // normalize input
            string value = text.Replace("_", "").Trim();

            // detect prefix
            string prefix;
            int numberBase;

            if (value.StartsWith("0b", StringComparison.OrdinalIgnoreCase) || value.StartsWith("b", StringComparison.OrdinalIgnoreCase))
            {
                prefix = value.StartsWith("0b", StringComparison.OrdinalIgnoreCase) ? value.Substring(0, 2) : value.Substring(0, 1);
                numberBase = 2;
            }
            else if (value.StartsWith("0x", StringComparison.OrdinalIgnoreCase) || value.StartsWith("x", StringComparison.OrdinalIgnoreCase) || value.StartsWith("h", StringComparison.OrdinalIgnoreCase))
            {
                prefix = value.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? value.Substring(0, 2) : value.Substring(0, 1);
                numberBase = 16;
            }
            else
            {
                return string.Empty;
            }

            string digits = value.Substring(prefix.Length);
            try
            {
                return Convert.ToInt64(digits, numberBase).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                if (numberBase == 16 && !text.StartsWith("0"))
                {
                    return "0" + text;
                }

                return text;
            }
        }


        public static List<string> ToLines(this string text)
        {
            if (text?.Length == 0)
            {
                return new List<string>() { string.Empty };
            }

            var lines = new List<string>();
            if (text != null)
            {
                using (var sr = new StringReader(text))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        lines.Add(line);
                    }
                }
            }

            return lines;
        }

        public static bool ContainsIgnoreCase(this string source, string toCheck)
        {
            if (source == null)
            {
                return true;
            }

            return source?.IndexOf(toCheck, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static string PadCenter(this string text, int length)
        {
            string line = "";
            int before = (length - text.Length) / 2;
            line += new string(' ', before);
            line += text;
            line = line.PadRight(length);
            return line;
        }

        public static bool IsOpened(this string filePath)
        {
            if (!File.Exists(filePath))
            {
                return false;
            }

            try
            {
                Stream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }


        public static string TrimSpace(this string input)
        {
            return input.Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "").Trim();
        }

        public static string SheetName2Block(this string name)
        {
            if (name.Equals("DCTEST_Func", StringComparison.OrdinalIgnoreCase))
            {
                return "IO";
            }

            if (name.Equals("DCTEST_IDCODE", StringComparison.OrdinalIgnoreCase))
            {
                return "JTAG";
            }

            return name;
        }

        public static string AddBlockFlag(this string source, string name)
        {
            if (string.IsNullOrEmpty(source))
            {
                return "F_" + name;
            }

            return source + ",F_" + name;
        }

        public static string PadBoth(this string source, int length)
        {
            int spaces = length - source.Length;
            int padLeft = spaces / 2 + source.Length;
            return source.PadLeft(padLeft).PadRight(length);
        }

        public static Dictionary<string, string> GetDict(this string line, char split1, char split2)
        {
            var dict = new Dictionary<string, string>();
            string[] arr = line.Split(new[] { split1 }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string item in arr)
            {
                string[] keyValue = item.Split(split2);
                if (keyValue.Length == 2)
                {
                    dict[keyValue[0]] = keyValue[1];
                }
            }
            return dict;
        }

        public static string GetSortPatNameForBinTable(this string patName)
        {
            if (string.IsNullOrEmpty(patName.Trim()))
            {
                return "";
            }

            string[] items = patName.Split('_');
            if (items.Length < 11)
            {
                return patName;
            }

            var resultList = new List<string>();
            resultList.Add(items[1]);
            resultList.Add(items[4]);
            resultList.Add(items[6]);
            resultList.Add(items[7]);
            resultList.Add(items[8]);
            if (items.Length >= 12)
            {
                resultList.Add(items[11]);
            }

            if (items.Length >= 13)
            {
                resultList.Add(items[12]);
            }

            string result = string.Join("_", resultList);
            return result.ToUpper();
        }


        public static string ReplaceDoubleSpace(this string text)
        {
            string result = text;
            do
            {
                result = result.Replace("  ", " ");
            } while (result.IndexOf("  ", StringComparison.Ordinal) >= 0);

            return result;
        }
    }
}
