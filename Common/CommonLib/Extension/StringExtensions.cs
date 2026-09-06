using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace CommonLib.Extension
{
    public static partial class StringExtensions
    {
        [GeneratedRegex(@"(\""[^""]+\"")")]
        private static partial Regex MyRegexQuotedSegment();
        [GeneratedRegex("^0x[a-f0-9]+$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegexStrictHex();
        [GeneratedRegex("^((0x)?|(x)?|(h)?)[a-f0-9]+$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegexHexCandidate();
        [GeneratedRegex("^((0b)?|(b)?)[0-7]+$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegexOctCandidate();

        public static string CalBitWidth(this string lsb, string msb)
        {
            int lsbBit = int.Parse(lsb);
            int msbBit = int.Parse(msb);
            int bitWidth = Math.Abs(msbBit - lsbBit) + 1;
            return bitWidth.ToString("");
        }

        public static bool IsHexOrOctValue(this string str)
        {
            if (MyRegexHexCandidate().IsMatch(str))
            {
                return true;
            }

            if (MyRegexOctCandidate().IsMatch(str))
            {
                return true;
            }

            return false;
        }

        public static bool IsHexValue(this string str)
        {
            return MyRegexStrictHex().IsMatch(str);
        }

        public static string EFuseBitToInt(this string value)
        {
            string dataType = "";
            string header = "";
            if (IsNumber(value))
            {
                return value;
            }

            value = value.Replace("_", "").Trim();
            List<string> lBitList = ["b", "B", "0b", "0B"];
            List<string> lHexList = ["x", "X", "0x", "0X", "h", "H"];

            foreach (string bitPat in lBitList)
            {
                if (value.Length >= bitPat.Length && bitPat == value[..bitPat.Length])
                {
                    header = bitPat;
                    dataType = "B";
                    break;
                }
            }

            foreach (string hexPat in lHexList)
            {
                if (value.Length >= hexPat.Length && hexPat == value[..hexPat.Length])
                {
                    header = hexPat;
                    dataType = "X";
                    break;
                }
            }

            if (string.IsNullOrEmpty(dataType))
            {
                return "";
            }

            string data = value.Replace(header, "");
            string result;
            switch (dataType)
            {
                case "B":
                    try
                    {
                        result = Convert.ToInt32(data, 2).ToString(CultureInfo.InvariantCulture);
                    }
                    catch (Exception)
                    {
                        result = value;
                    }
                    break;
                case "X":
                    try
                    {
                        result = Convert.ToInt32(data, 16).ToString(CultureInfo.InvariantCulture);
                    }
                    catch (Exception)
                    {
                        result = value;
                        if (value[..1] != "0")
                        {
                            result = "0" + value;
                        }
                    }
                    break;
                default:
                    result = "";
                    break;
            }

            return result;
        }

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

            if (value.StartsWithIgnoreCase("0b") || value.StartsWithIgnoreCase("b"))
            {
                prefix = value.StartsWithIgnoreCase("0b") ? value[..2] : value[..1];
                numberBase = 2;
            }
            else if (value.StartsWithIgnoreCase("0x") || value.StartsWithIgnoreCase("x") || value.StartsWithIgnoreCase("h"))
            {
                prefix = value.StartsWithIgnoreCase("0x") ? value[..2] : value[..1];
                numberBase = 16;
            }
            else
            {
                return string.Empty;
            }

            string digits = value[prefix.Length..];
            try
            {
                return Convert.ToInt64(digits, numberBase).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                if (numberBase == 16 && !text.StartsWith('0'))
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
                return [string.Empty];
            }

            var lines = new List<string>();
            if (text != null)
            {
                using var sr = new StringReader(text);
                string? line;
                while ((line = sr.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }

            return lines;
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
                FileStream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
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
            int padLeft = (spaces / 2) + source.Length;
            return source.PadLeft(padLeft).PadRight(length);
        }

        public static Dictionary<string, string> GetDict(this string line, char split1, char split2)
        {
            var dict = new Dictionary<string, string>();
            string[] arr = line.Split([split1], StringSplitOptions.RemoveEmptyEntries);
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

            var resultList = new List<string>
            {
                items[1],
                items[4],
                items[6],
                items[7],
                items[8]
            };
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
            } while (result.Contains("  "));

            return result;
        }

        public static string ToExcelColName(this string startCol, int offset)
        {
            if (string.IsNullOrEmpty(startCol))
            {
                return string.Empty;
            }

            int index = 0;
            foreach (char c in startCol.ToUpper())
            {
                if (c < 'A' || c > 'Z')
                {
                    return string.Empty;
                }
                index = (index * 26) + c - 'A' + 1;
            }

            int targetIndex = index - 1 + offset;

            string columnName = string.Empty;
            while (targetIndex >= 0)
            {
                columnName = (char)('A' + (targetIndex % 26)) + columnName;
                targetIndex = (targetIndex / 26) - 1;
            }
            return columnName;
        }

        public static List<List<string>> CsvConvertToLists(this string file)
        {
            string keyComma = "#comma>";
            string keyNewline = "#newline>";
            string keyDouble = "#double>";
            string text = File.ReadAllText(file).Replace("\"\"", keyDouble);
            string[] matches = MyRegexQuotedSegment().Split(text);
            var sb = new StringBuilder();
            foreach (string match in matches)
            {
                if (match.StartsWith('\"') && match.EndsWith('\"'))
                {
                    sb.Append(match.Trim('"').Replace(",", keyComma).Replace("\r\n", keyNewline).Replace("\n", keyNewline));
                }
                else
                {
                    sb.Append(match);
                }
            }

            List<string> lines = sb.ToString().ToLines();
            var lists = new List<List<string>>(lines.Count);
            foreach (string line in lines)
            {
                string[] cells = line.Split(',');
                var row = new List<string>(cells.Length);
                foreach (string cell in cells)
                {
                    row.Add(cell.Replace(keyComma, ",").Replace(keyNewline, "\n").Replace(keyDouble, "\""));
                }
                lists.Add(row);
            }

            return lists;
        }
    }
}
