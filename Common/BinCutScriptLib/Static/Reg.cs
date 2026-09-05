using System.Text.RegularExpressions;

using TestPlanLib.BinCut.Flow;

namespace BinCutScriptLib.Static
{
    public static partial class Reg
    {
        public static readonly Regex RegexAssignment = MyRegex();

        public static readonly Regex RegexUnit = MyRegex1();
        public static readonly Regex RegexDoAll = MyRegex2();

        public static readonly Regex RegexValue1 = MyRegex3();
        public static readonly Regex RegexPin = MyRegex4();

        public static readonly Regex RegexExpress = MyRegex5();
        public static readonly Regex RegexMode = MyRegex6();

        public static readonly Regex RegexBincitVbtConfigValue = MyRegex7();
        public static readonly Regex RegexBincitConfigValue = MyRegex8();
        public static readonly Regex RegexHarvFlags = MyRegex9();
        public static readonly Regex RegexHarvResult = MyRegex10();
        public static readonly Regex RegexStartEqn = MyRegex11();

        public static readonly Regex RegexLvCoreVoltage = MyRegex12();
        public static readonly Regex RegexLvCoreEvaluate = MyRegex13();
        public static readonly Regex RegexLvCoreResult = MyRegex14();
        public static readonly Regex RegexCoreHvcc = MyRegex15();
        public static readonly Regex RegexModeBinProductGb = MyRegex16();
        public static readonly Regex RegexModeBinProduct = MyRegex17();
        public static readonly Regex RegexCoreE1ProductGb = MyRegex18();
        public static readonly Regex RegexCoreE1Product = MyRegex19();
        public static readonly Regex RegexmVWithBinSearchMode = MyRegex20();
        public static readonly Regex RegexmVWithProductMode = MyRegex21();
        public static readonly Regex RegexAllProduct = MyRegex22();
        public static readonly Regex RegexAllmVWithMode = MyRegex23();
        public static readonly Regex RegexAllmV = MyRegex24();
        public static readonly Regex RegexAllmV1 = MyRegex25();
        public static readonly Regex RegexAllmV2 = MyRegex26();
        public static readonly Regex RegexAllmV3 = MyRegex27();
        public static readonly Regex RegexBinningVmax = MyRegex28();
        public static readonly Regex RegexCpVmax = MyRegex29();

        public static readonly Regex RegexAllRatio = MyRegex30();

        public static readonly Regex RegexFunction = MyRegex31();

        public static readonly Regex RegexRegexPerformance = MyRegex32();

        public static readonly Regex RegexcsharpMatchFlag = MyRegex33();
        public static readonly Regex RegexcsharpSimpleMatchFlag = MyRegex34();
        public static readonly Regex RegexvbtMatchFlag = MyRegex35();

        public static readonly Regex RegexcleanLine = MyRegex36();

        public static readonly Regex RegexSite1 = MyRegex37();
        public static readonly Regex RegexSite2 = MyRegex38();
        public static readonly Regex RegexSite3 = MyRegex39();

        public static readonly Regex RegexType = MyRegex40();

        public static readonly Regex RegexPerformance = MyRegex41();

        public static readonly Regex RegexSelSram = MyRegex42();
        public static readonly Regex RegexSelSram1 = MyRegex43();
        public static readonly Regex RegexSelSram2 = MyRegex44();
        public static readonly Regex RegexDigSrcAssignment = MyRegex45();

        public static readonly Regex RegexLsb = MyRegex46();

        public static readonly Regex RegexVoltage = MyRegex47();

        public static readonly Regex RegexB = MyRegex48();
        public static readonly Regex RegexValue = MyRegex49();
        //eg.  23.x211h

        public static readonly Regex RegxChannel = MyRegex50();

        public static readonly Regex RegexLine = MyRegex51();
        public static readonly Regex RegexLine1 = MyRegex52();

        //[INFO]  [Site 1] Assign read back value '0' for fuse 'EMA_0000308' to cache.
        public static readonly Regex RegexAssignFuse = MyRegex53();

        //Site: 0, VDD_SOC_SRAM_OFFSET Value: 3, VDD_SOC_SRAM = 678
        public static readonly Regex Regexoffset = MyRegex54();
        //=> Site: 2, VDD_SOC_SRAM OFFSET: 0.003V
        public static readonly Regex RegexOldRegexoffset = MyRegex55();

        public static readonly Regex RegexBv = MyRegex56();
        public static readonly Regex RegexDssc = MyRegex57();

        //Dssc, Selsram
        public static readonly Regex RegexHarvestSourceCodeAssignment = MyRegex58();
        public static readonly Regex RegexSelsramAssignment = MyRegex59();
        public static readonly Regex RegexHarvestSourceCodeMatchRow = MyRegex60();

        public static readonly Regex RegexFuse1 = MyRegex61();
        public static readonly Regex RegexFuse2 = MyRegex62();

        public static readonly Regex RegexC = MyRegex63();
        public static readonly Regex RegexModNum = MyRegex64();
        public static readonly Regex RegxValue1 = MyRegex65();
        public static readonly Regex RegxValue2 = MyRegex66();
        public static readonly Regex RegexFormula = MyRegex67();

        public static readonly Regex RegexSplit = MyRegex68();
        public static readonly Regex RegexSplit1 = MyRegex69();

        public static readonly Regex RegContainPerformanceModeWithGroup = MyRegex70();

        [GeneratedRegex(@"\((.*?)\)", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        [GeneratedRegex("[a-zA-Z]*", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"\bDoAll\b\s*[:]\s*[']?(True|False)[']?", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(@"\d+\.\d{2}", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex("^VDD_GPU[0-9]", RegexOptions.Compiled)]
        private static partial Regex MyRegex4();
        [GeneratedRegex(@"(?<value>\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex5();
        [GeneratedRegex(@"^M{1}[A-Z]+\w{3}$", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex6();
        [GeneratedRegex(@"^(\w+.*)=", RegexOptions.Compiled)]
        private static partial Regex MyRegex7();
        [GeneratedRegex(@"^\[INFO\]\s+(.*?):\s*(.*)", RegexOptions.Compiled)]
        private static partial Regex MyRegex8();
        [GeneratedRegex(@"\[Site (\d+)\] Flag '(.*?)' = (.*)", RegexOptions.Compiled)]
        private static partial Regex MyRegex9();
        [GeneratedRegex(@"^(site:\d+.*)", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex10();
        [GeneratedRegex(@"\[INFO]\s*Start_EQN:\s*mode=(.*?),\s*eqn=(.*)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex11();
        [GeneratedRegex(BinCutFlowTable.RegexLvCoreVoltage, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex12();
        [GeneratedRegex(BinCutFlowTable.RegexLvCoreEvaluate, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex13();
        [GeneratedRegex(BinCutFlowTable.RegexLvCoreResult, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex14();
        [GeneratedRegex(BinCutFlowTable.RegexCoreHvcc, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex15();
        [GeneratedRegex(BinCutFlowTable.RegexModeBinProductGb, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex16();
        [GeneratedRegex(BinCutFlowTable.RegexModeBinProduct, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex17();
        [GeneratedRegex(BinCutFlowTable.RegexCoreE1ProductGb, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex18();
        [GeneratedRegex(BinCutFlowTable.RegexCoreE1Product, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex19();
        [GeneratedRegex(BinCutFlowTable.RegexmVWithBinSearchMode, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex20();
        [GeneratedRegex(BinCutFlowTable.RegexmVWithProductMode, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex21();
        [GeneratedRegex(BinCutFlowTable.RegexAllProduct, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex22();
        [GeneratedRegex(BinCutFlowTable.RegexAllmVWithMode, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex23();
        [GeneratedRegex(BinCutFlowTable.RegexAllmV, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex24();
        [GeneratedRegex(@"^(?<value>[+|-]?\s*\d+(\.\d+)?)\s*mV", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex25();
        [GeneratedRegex(@"^(?<value>[+|-]?\s*\d+(\.\d+)?)\s*mV$", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex26();
        [GeneratedRegex(@"(?<value>[+|-]?\s*\d+(\.\d+)?)\s*mV$", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex27();
        [GeneratedRegex(BinCutFlowTable.RegexBinningVmax, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex28();
        [GeneratedRegex(BinCutFlowTable.RegexCpVmax, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex29();
        [GeneratedRegex(BinCutFlowTable.RegexAllRatio, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex30();
        [GeneratedRegex(BinCutFlowTable.RegexFunction, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex31();
        [GeneratedRegex("(?<pmode>M[a-zA-Z0-9]{4}[a-zA-Z0-9]?)", RegexOptions.Compiled)]
        private static partial Regex MyRegex32();
        [GeneratedRegex(@"(\[*\w*\]*\s*\w+\s*:\s*'[^']+\s*'\.*)", RegexOptions.Compiled)]
        private static partial Regex MyRegex33();
        [GeneratedRegex(@"(\[INFO]\s*)", RegexOptions.Compiled)]
        private static partial Regex MyRegex34();
        [GeneratedRegex(@"Active EnableWords\s*:\s*\w+\s*", RegexOptions.Compiled)]
        private static partial Regex MyRegex35();
        [GeneratedRegex("(?<pmode>_M[a-zA-Z0-9]{4}[a-zA-Z0-9]?)", RegexOptions.Compiled)]
        private static partial Regex MyRegex36();
        [GeneratedRegex(@"\[Site\s*(?<site>\d+)\]", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex37();
        [GeneratedRegex(@"\(|\)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex38();
        [GeneratedRegex(@"Site\((?<site>\d+)\)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex39();
        [GeneratedRegex(@"[\(](?<type>[^)]+)[\)]", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex40();
        [GeneratedRegex("(?<pmode>M[a-zA-Z0-9]{4}[a-zA-Z0-9]?)", RegexOptions.Compiled)]
        private static partial Regex MyRegex41();
        [GeneratedRegex(@"\[Site (?<site>\d+)\]\s+(?<bitString>\w+)\(\w+\)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex42();
        [GeneratedRegex(@"^\[INFO\]\s+\[Site \d+\]\s+\d+$", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex43();
        [GeneratedRegex(@"^\[INFO\]\s+\[Site \d+\]\s+\d+(\(Selsram\(.*\)\)|\(DSSC\(.*\)\))", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex44();
        [GeneratedRegex(@"\(([^)]*)\)")]
        private static partial Regex MyRegex45();
        [GeneratedRegex(@"\(LSB->MSB\)", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex46();
        [GeneratedRegex(@"(?<value>[+|-]*[\d|.]+)\s*mV", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex47();
        [GeneratedRegex("^b", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex48();
        [GeneratedRegex(@"\d+(\.\d+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex49();
        [GeneratedRegex(@"\d+\.[a-z]\d+\w*", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex50();
        [GeneratedRegex(@"\d+\s+\d+", RegexOptions.Compiled)]
        private static partial Regex MyRegex51();
        [GeneratedRegex(@"\s+Number\s+Site\s+Test\s+Name\s+Pin", RegexOptions.Compiled)]
        private static partial Regex MyRegex52();
        [GeneratedRegex(@"\[Site (?<site>\d+)\].*?value '(?<value>\d+)' for fuse '(?<pmode>\w+)'", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex53();
        [GeneratedRegex(@"OFFSET .*:\s?(?<offset>-?(\d+[.])?\d+)", RegexOptions.Compiled)]
        private static partial Regex MyRegex54();
        [GeneratedRegex(@"OFFSET:\s?(?<offset>-?(\d+[.])?\d+)", RegexOptions.Compiled)]
        private static partial Regex MyRegex55();
        [GeneratedRegex("^BV_", RegexOptions.Compiled)]
        private static partial Regex MyRegex56();
        [GeneratedRegex(@"\((?<site>\d+)\).*Read from DSSC :[\s_]*(?<Name>\w+)\s+.*=\s*(?<value1>.+)\s*\[(?<value2>\d+)\]", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex57();
        [GeneratedRegex("^[A-Za-z]=Dssc", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex58();
        [GeneratedRegex("^[A-Za-z]=Selsram", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex59();
        [GeneratedRegex(@"^\[INFO\]\s+\[Site \d+\]\s+\d+(\(DSSC\(\w*\)\))", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex60();
        [GeneratedRegex(@"\[Site (?<site>\d+)\] Set fuse value in cache '(?<name>\w+)' = 'SiteGeneric`1 \{(?<value>.*?)\}'", RegexOptions.None)]
        private static partial Regex MyRegex61();
        [GeneratedRegex(@"\[Site (?<site>\d+)\] Set fuse value in cache '(?<name>\w+)' =\s*(?<value>\d+)\s*", RegexOptions.None)]
        private static partial Regex MyRegex62();
        [GeneratedRegex("_C$", RegexOptions.None)]
        private static partial Regex MyRegex63();
        [GeneratedRegex(@"^[a-zA-z]+(?<modeNum>[0-9]+\w)", RegexOptions.Compiled)]
        private static partial Regex MyRegex64();
        [GeneratedRegex(@"\d+\.\d{4}", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex65();
        [GeneratedRegex(@"\d+\.\w{2}", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex66();
        [GeneratedRegex(@"(?<Name>\w+)", RegexOptions.Compiled)]
        private static partial Regex MyRegex67();
        [GeneratedRegex("([&|()])", RegexOptions.Compiled)]
        private static partial Regex MyRegex68();
        [GeneratedRegex("([+|-|*|/])", RegexOptions.Compiled)]
        private static partial Regex MyRegex69();
        [GeneratedRegex("(?!Mbist)(?<pmode>M([a-zA-Z]){1}([a-zA-Z0-9]){1}(?<modenumber>[a-fA-F0-9|x|X]{3}))", RegexOptions.Compiled)]
        private static partial Regex MyRegex70();
    }
}
