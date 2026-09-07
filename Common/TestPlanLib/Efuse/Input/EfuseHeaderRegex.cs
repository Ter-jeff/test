using System.Text.RegularExpressions;

namespace TestPlanLib.Efuse.Input
{
    internal static partial class EfuseHeaderRegex
    {
        private const string ConTestName = "^Test Name";
        private const string ConType = "^Type";
        private const string ConBank = "^Bank";
        private const string ConWriteRead = "^Write/Read";
        private const string ConPurpose = "^Purpose";
        private const string ConUserdefined = "^User_defined";
        private const string ConInitPat = "^Init pattern name";
        private const string ConPayloadPat = "^PL pattern name";
        private const string ConJob = "CP1|CP2|WLFT1|FT1|WLFT2|FT2|FT3|SLT";

        [GeneratedRegex(ConTestName, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex TestName();
        [GeneratedRegex(ConType, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex Type();
        [GeneratedRegex(ConBank, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex Bank();
        [GeneratedRegex(ConWriteRead, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex WriteRead();
        [GeneratedRegex(ConPurpose, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex Purpose();
        [GeneratedRegex(ConUserdefined, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex UserDefined();
        [GeneratedRegex(ConInitPat, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex InitPat();
        [GeneratedRegex(ConPayloadPat, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex PayloadPat();
        [GeneratedRegex(ConJob, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex Job();
    }
}
