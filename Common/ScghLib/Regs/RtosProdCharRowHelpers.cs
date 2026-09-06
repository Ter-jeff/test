using System.Text.RegularExpressions;

namespace ScghLib.Reader
{
    internal static partial class RtosProdCharRowHelpers
    {
        [GeneratedRegex("0x(?<PmgrMode>[0-9A-Fa-f]+)", RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex();
        [GeneratedRegex(@"(?<Scenario>\d+\s+)", RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex1();
    }
}
