
namespace Cautogen.AutoCZ.CharPreProcessor.InputDataBase
{
    public class InputDefRow
    {
        public string BlockName { get; set; }
        public string Group { get; set; }
        public string DcCategory { get; set; }
        public string AcCategory { get; set; }
        public string LevelSheet { get; set; }
        public string TimingSheet { get; set; }
        public string PowerRunScenario { get; set; }
        public bool UsePMode { get; set; }

        public InputDefRow()
        {
            BlockName = "";
            Group = "";
            DcCategory = "";
            AcCategory = "";
            LevelSheet = "";
            TimingSheet = "";
            PowerRunScenario = "";
            UsePMode = false;
        }
    }
}
