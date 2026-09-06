using System.Text.RegularExpressions;

using CommonLib.Enums;

namespace CommonLib.Extension
{
    public static partial class StringExtensions
    {
        [GeneratedRegex(@"_IN\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexInitMarker();
        [GeneratedRegex(@"_PL\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexPayloadMarker();

        private static readonly Regex _regexInitMarker = MyRegexInitMarker();
        private static readonly Regex _regexPayloadMarker = MyRegexPayloadMarker();

        public static EnumPatternType GetPatternType(this string pattern)
        {
            if (pattern.ContainsIgnoreCase("_RTS_"))
            {
                return EnumPatternType.RTOS;
            }

            if (pattern.ContainsIgnoreCase("_AN_"))
            {
                return EnumPatternType.HARDIP;
            }

            if (_regexInitMarker.IsMatch(pattern))
            {
                return EnumPatternType.Init;
            }

            if (_regexPayloadMarker.IsMatch(pattern) || pattern.ContainsIgnoreCase("_FULP_"))
            {
                return EnumPatternType.Payload;
            }

            if (pattern.StartsWithIgnoreCase("RETENTION_PAUSE_"))
            {
                return EnumPatternType.RetentionWait;
            }

            return EnumPatternType.Unknown;
        }

        public static bool IsInit(this string pattern)
        {
            if (_regexInitMarker.IsMatch(pattern))
            {
                return true;
            }

            return false;
        }
    }
}
