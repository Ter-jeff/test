using System.Text.RegularExpressions;

namespace TagDiffCore.Regs
{
    internal static partial class Reg
    {
        [GeneratedRegex("Selector[s]?", RegexOptions.IgnoreCase)]
        public static partial Regex RegexSelector();

        [GeneratedRegex("^H|HARDIP|LCD|DCTEST", RegexOptions.IgnoreCase)]
        public static partial Regex RegexHardip();
    }
}
