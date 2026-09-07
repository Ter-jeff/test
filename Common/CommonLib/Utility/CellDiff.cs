using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace CommonLib.Utility
{
    public static class CellDiff
    {
        public static bool IsLiked(string text, string patten)
        {
            if (patten.Contains(".*") ||
                patten.Contains(".+") ||
                patten.Contains('|'))
            {
                return Regex.IsMatch(FormatStringForCompare(text), "^" + FormatStringForCompare(patten) + "$");
            }

            return FormatStringForCompare(text).EqualsIgnoreCase(patten);
        }

        public static string FormatStringForCompare(string text)
        {
            string result = text.Trim();

            result = ReplaceDoubleBlank(result);

            result = result.ToUpper();

            return result;
        }

        private static string ReplaceDoubleBlank(string text)
        {
            string result = text;
            do
            {
                result = result.Replace("  ", " ");
            } while (result.Contains("  "));

            return result;
        }
    }
}
