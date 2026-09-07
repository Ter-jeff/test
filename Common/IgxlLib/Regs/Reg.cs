using System.Text.RegularExpressions;

namespace IgxlLib.Regs
{
    public static partial class Reg
    {
        public static readonly Regex InitRegex = CreateInitRegex();
        public static readonly Regex WhitespaceRegex = CreateWhitespaceRegex();

        [GeneratedRegex(@"_IN\w{2}_", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex CreateInitRegex();
        [GeneratedRegex(@"\s+", RegexOptions.Compiled)]
        private static partial Regex CreateWhitespaceRegex();
    }
}
