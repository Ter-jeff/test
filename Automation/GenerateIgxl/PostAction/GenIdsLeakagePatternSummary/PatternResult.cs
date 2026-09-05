using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace CommonLib.IdsLeakageCell
{
    [ExcludeFromCodeCoverage]
    public class PatternResult
    {
        public string Pattern { get; set; }
        public string PatternVersion { get; set; }
        public List<string> Xpins { get; set; }
        public List<string> UnusedIoPins { get; set; }
        public List<string> PatExtraPins { get; set; }
        public List<string> MissSubFunction { get; set; }
    }
}
