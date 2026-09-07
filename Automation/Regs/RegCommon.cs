using System.Text.RegularExpressions;

namespace Automation.Regs
{
    internal class RegCommon
    {
        internal static Regex Flow { get; } = new Regex("^Flow_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        internal static Regex Regex { get; } = new Regex("^F", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        internal static Regex RegexInit { get; } = new Regex(@"^init(?<idx>\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }
}
