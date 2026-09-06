using System.Text.RegularExpressions;

namespace TestPlanLib.Static
{
    internal static partial class Reg
    {
        internal static readonly Regex _regex = MyRegex();
        internal static readonly Regex _regex2 = MyRegex1();
        internal static readonly Regex _regex3 = MyRegex3();
        internal static readonly Regex _regex4 = MyRegex2();

        [GeneratedRegex(@"\s{2,}", RegexOptions.Compiled)]
        private static partial Regex MyRegex();

        [GeneratedRegex(@"^M[a-zA-Z]+\d+$", RegexOptions.Compiled)]
        private static partial Regex MyRegex2();
        [GeneratedRegex("^(HV|LV|NV)$", RegexOptions.Compiled)]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"\w+:\w+\((?<Flag>\w+)\)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex3();
    }
}
