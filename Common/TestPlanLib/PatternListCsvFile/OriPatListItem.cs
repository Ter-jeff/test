namespace TestPlanLib.PatternListCsvFile
{
    public class OriPatListItem
    {
        public int RowNum { get; set; } = 1;
        //Both
        public string Idx { get; set; } = string.Empty;
        //  Both
        public string Pattern { get; set; } = string.Empty;
        //Only TW
        public string LatestVersion { get; set; } = string.Empty;
        //If not exist TW will added
        public string ReleaseDate { get; set; } = string.Empty;
        //Both
        public string UseNoUse { get; set; } = string.Empty;
        //If not exist TW will added
        public string DRi { get; set; } = string.Empty;
        //If not exist TW will added
        public string ReleaseNote { get; set; } = string.Empty;
        //If not exist TW will added
        public string RadarNum { get; set; } = string.Empty;
        //If not exist TW will added
        public string Org { get; set; } = string.Empty;
        //If not exist TW will added
        public string TypeSpec { get; set; } = string.Empty;

        public string TimeSetLatest //Ori pattern List
        {
            get;
            set;
        } = "";

        public string TimeSetVersion           //TW Pattern List
        {
            get { return TimeSetLatest; }
            set { TimeSetLatest = value; }
        }
        //Both

        public string FileVersions { get; set; } = string.Empty;
        //Added in TW
        public string OpCode { get; set; } = string.Empty;
        //Added in TW
        public string ScanMode { get; set; } = string.Empty;
        //Added in TW
        public string Halt { get; set; } = string.Empty;
        //Added in TW
        public string Compilation { get; set; } = string.Empty;
        //Added in TW
        public string HLv { get; set; } = string.Empty;
        //Added in TW
        public string OriTimeMod { get; set; } = string.Empty;
        //Added in TW
        public string CheckRt { get; set; } = string.Empty;
        //Added in TW
        public string TpCategory { get; set; } = string.Empty;
        //Added in TW
        public string ScanTSet { get; set; } = string.Empty;
    }

}
