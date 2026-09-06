using System.Text.RegularExpressions;

namespace ScghLib.Utility
{
    internal static partial class BistActionHelpers
    {
        // Get flag value
        public const string ConKeyWordGet = @"GET[_](?<str>.*)|SRV[EP]\d*_GET[_](?<str>.*)";
        // Get flag value
        public const string ConKeyWordMultiGet = @"GET[ ]\((?<str>.*)\)";
        public const string ConKeyWordCheck = @"CHECK[_](?<str>.*)|SRV[EP]\d*_CHECK[_](?<str>.*)";
        //SET_FINAL_TEST_0
        public const string ConKeyWordSet = @"SET[_](?<str>.*)[_](?<str1>\d*)$|SRV[EP]\d*_SET[_](?<str>.*)[_](?<str1>\d*)$";
        //SET_DIE_TYPE_REPAIRED
        public const string ConKeyWordSetDefault = @"SET[_](?<str>.*)$|SRV[EP]\d*_SET[_](?<str>.*)$";
        public const string ConKeyWordFail = @"^FAIL$|SRV[EP]\d*_FAIL$";
        public const string ConKeyWordFailCheckScan = "FAIL_SCAN_PF_CHECK$";
        public const string ConKeyWordFailCheckScan1 = @"CHECK_SCAN_\w+_PASS";
        public const string ConKeyWordFailCheck = "FAIL_CHECK_(?<str>.*)$";
        public const string ConKeyWordPass = @"^PASS$|^SRV[EP]\d*_PASS$";
        //|^SRV[EP]\d*_(INT)*RETENTION_PAUSE.*$";
        public const string ConKeyWordRetention = ".*RETENTION(_BIR)?_PAUSE.*";
        //"|^SRV[EP]\d*_(INT)*RETENTION_VDD.*$";
        public const string ConKeyWordRetentionVoltDrop = ".*RETENTION(_BIR)?_VDD.*";
        public const string ConKeyWordLoopStart = @"LOOP_START_(?<str>\d+)";
        public const string ConKeyWordLoopEnd = "LOOP_END";
        public const string ConKeyWordDomainStart = @"DOMAIN_START_(?<str>\w+)";
        public const string ConKeyWordDomainEnd = @"DOMAIN_END_(?<str>\w+)";
        public const string ConKeyWordCallSubFlow = "CALLSUBFLOW_.*";
        public const string ConKeyWordSetSiteVar = @"SETVAR[_](?<str>.*)[_](?<str1>\d*)$|SRV[EP]\d*_SET[_](?<str>.*)[_](?<str1>\d*)$";

        [GeneratedRegex(@".*RETENTION(_BIR)?_VDD_STEP(?<str>\d+)", RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex();
        [GeneratedRegex(@".*RETENTION(_BIR)?_PAUSE_(?<str>\d+)", RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex1();
        [GeneratedRegex(ConKeyWordDomainStart, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex10();
        [GeneratedRegex(ConKeyWordDomainStart, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex11();
        [GeneratedRegex(ConKeyWordRetention, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex12();
        [GeneratedRegex(ConKeyWordRetentionVoltDrop, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex13();
        [GeneratedRegex(ConKeyWordCheck, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex14();
        [GeneratedRegex(ConKeyWordLoopStart, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex15();
        [GeneratedRegex(ConKeyWordLoopEnd, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex16();
        [GeneratedRegex(ConKeyWordDomainStart, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex17();
        [GeneratedRegex(ConKeyWordDomainEnd, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex18();
        [GeneratedRegex(ConKeyWordFailCheck, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex19();
        [GeneratedRegex(ConKeyWordSetDefault, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex2();
        [GeneratedRegex(ConKeyWordFailCheckScan, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex20();
        [GeneratedRegex(ConKeyWordFailCheckScan1, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex21();
        [GeneratedRegex(ConKeyWordFail, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex22();
        [GeneratedRegex(ConKeyWordGet, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex23();
        [GeneratedRegex(ConKeyWordMultiGet, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex24();
        [GeneratedRegex(ConKeyWordPass, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex25();
        [GeneratedRegex(ConKeyWordSet, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex26();
        [GeneratedRegex(ConKeyWordSetDefault, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex27();
        [GeneratedRegex(ConKeyWordRetention, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex28();
        [GeneratedRegex(ConKeyWordRetentionVoltDrop, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex29();
        [GeneratedRegex(ConKeyWordSet, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex3();
        [GeneratedRegex(ConKeyWordCallSubFlow, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex30();
        [GeneratedRegex(ConKeyWordSetSiteVar, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex31();
        [GeneratedRegex(ConKeyWordSet, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex4();
        [GeneratedRegex(ConKeyWordGet, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex5();
        [GeneratedRegex(@".*RETENTION(_BIR)?_PAUSE_(?<str>\d+)", RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex6();
        [GeneratedRegex(ConKeyWordMultiGet, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex7();
        [GeneratedRegex(ConKeyWordLoopStart, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex8();
        [GeneratedRegex(ConKeyWordLoopStart, RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex MyRegex9();
    }
}
