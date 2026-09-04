using System;
using System.IO;
using System.Text.RegularExpressions;

namespace CommonLib.Utility.StringExtension
{
    public static class Pattern
    {
        private static readonly Regex _regex = new Regex(@"_RTS_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(@"_AN_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex3 = new Regex(@"_IN\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex4 = new Regex(@"_PL\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex5 = new Regex(@"_IN\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex6 = new Regex(@"_PL\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex7 = new Regex(@"_PL\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex8 = new Regex(@"_IN\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex9 = new Regex(@"_FULP_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regexRetenWait = new Regex(@"^RETENTION_PAUSE_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _validNameRegex = new Regex("^((dd_)|(cz_)|(pp_)|(mn_)|(ht_)).*", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _instanceNameRegex = new Regex(@"^Instance[\s]*:[\s]*(?<InsName>[\w]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool IsValidPatName(this string patternName)
        {
            return _validNameRegex.IsMatch(patternName) || patternName.Equals("No_patt", StringComparison.OrdinalIgnoreCase) ||
                   _instanceNameRegex.IsMatch(patternName);
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

        public static EnmPatternType GetPatternType(this string pattern)
        {
            if (_regex.IsMatch(pattern))
            {
                return EnmPatternType.RTOS;
            }

            if (_regex2.IsMatch(pattern))
            {
                return EnmPatternType.HARDIP;
            }

            if (_regex3.IsMatch(pattern))
            {
                return EnmPatternType.Init;
            }

            if (_regex4.IsMatch(pattern) || _regex9.IsMatch(pattern))
            {
                return EnmPatternType.Payload;
            }

            if (_regexRetenWait.IsMatch(pattern))
            {
                return EnmPatternType.RetentionWait;
            }

            return EnmPatternType.Unknow;
        }

        public static EnmPatternType GetInOrPlPatternType(this string pattern)
        {
            if (_regex5.IsMatch(pattern))
            {
                return EnmPatternType.Init;
            }

            if (_regex6.IsMatch(pattern))
            {
                return EnmPatternType.Payload;
            }

            return EnmPatternType.Unknow;
        }

        public static bool IsPayLoad(this string pattern)
        {
            if (_regex7.IsMatch(pattern))
            {
                return true;
            }

            return false;
        }

        public static bool IsInit(this string pattern)
        {
            if (_regex8.IsMatch(pattern))
            {
                return true;
            }

            return false;
        }
    }

    public enum EnmPatternType
    {
        Init, RTOS, HARDIP, Payload, RetentionWait, Unknow

    }
}
