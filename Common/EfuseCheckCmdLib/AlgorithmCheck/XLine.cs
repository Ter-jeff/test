using System;
using System.Text.RegularExpressions;

namespace EfuseCheckCmdLib.AlgorithmCheck
{
    public partial class XLine
    {
        public static readonly Regex RegexSite1 = Site1Regex();
        public static readonly Regex RegexSite3 = Site3Regex();

        [GeneratedRegex(@"\[Site\s*(?<site>\d+)\]", RegexOptions.IgnoreCase)]
        private static partial Regex Site1Regex();

        [GeneratedRegex(@"Site\((?<site>\d+)\)", RegexOptions.IgnoreCase)]
        private static partial Regex Site3Regex();
        public LineType Type = LineType.Undefine;
        public int LineNo;
        public string Line = "";

        internal int GetSite3()
        {
            if (int.TryParse(RegexSite3.Match(Line).Groups["site"].ToString(), out int site))
            {
                return site;
            }

            return -1;
        }

        internal int GetSite()
        {
            if (int.TryParse(RegexSite1.Match(Line).Groups["site"].ToString(), out int site))
            {
                return site;
            }

            return -1;
        }

        public static long HexToDecimal(string hex)
        {
            return Convert.ToInt64(hex, 16);
        }

        public static long BinaryToDecimal(string binary)
        {
            return Convert.ToInt64(binary, 2);
        }
    }
}
