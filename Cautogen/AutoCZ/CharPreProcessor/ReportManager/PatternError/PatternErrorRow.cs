
namespace Cautogen.AutoCZ.CharPreProcessor.ReportManager.PatternError
{
    public class PatternErrorRow
    {
        public string Name { get; set; }
        public string SheetName { get; set; }
        public string WrongMeasCount { get; set; }
        public string WrongMeasOrder { get; set; }
        public string MissingMeasPin { get; set; }
        public string MissingPinSeq { get; set; }
        public string MissingPatternInPatInfo { get; set; }
        public string MissingPatternInPatList { get; set; }
        public string PatternDontUseInPatList { get; set; }
        public string PatternFileVersionNotMatch { get; set; }
        public string VersionInPatList { get; set; }
        public string PatternDontExistOnServer { get; set; }
        public string PatternCompileSuccess { get; set; }
        public string LatestVersionInServer { get; set; }

        public PatternErrorRow(string name)
        {
            Name = name;
        }
    }
}
