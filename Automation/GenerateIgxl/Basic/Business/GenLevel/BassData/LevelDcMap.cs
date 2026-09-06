namespace Automation.GenerateIgxl.Basic.Business.GenLevel.BassData
{
    public class LevelDcMap
    {
        public string LevelSheetName { set; get; }
        public string DcSpecName { set; get; }

        public LevelDcMap(string levelName, string dcSpecName)
        {
            LevelSheetName = levelName;
            DcSpecName = dcSpecName;
        }
    }
}
