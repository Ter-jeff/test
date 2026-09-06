namespace TestPlanLib.Basic
{
    public class PatternData
    {
        public string PatternName { get; set; } = string.Empty;
        public int RowNum { get; set; } = 1;
        public string LatestVersion { get; set; } = string.Empty;
        public string Use { get; set; } = string.Empty;
        public string Org { get; set; } = string.Empty;
        public string TypeSpec { get; set; } = string.Empty;
        public string TimeSetVersion { get; set; } = string.Empty;
        public string FileVersion { get; set; } = string.Empty;
        public string OpCode { get; set; } = string.Empty;
        public string ScanMode { get; set; } = string.Empty;
        public string Halt { get; set; } = string.Empty;
        public string OriginalTimingMode { get; set; } = string.Empty;
        public string Check { get; set; } = string.Empty;
        public string TpCategory { get; set; } = string.Empty;
        public string CheckComment { get; set; } = string.Empty;
        public string PatternVersion { get; set; } = string.Empty;
        public string ScanTset { get; set; } = string.Empty;
        public bool IsExist { get; set; }
        public string TimesetLatest { get; set; } = string.Empty;
    }
}
