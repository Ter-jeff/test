using RF_PatternTool;

using RfLib.Dvdc.Reader.CapturePostProcess;

namespace RFPatternTool
{
    public class PatternGenItem
    {
        public string Pattern;
        public List<string> InitDictionary;
        public KeyValuePair<string, string> Patset;
        public List<KeyValuePair<string, string>> PatSubr = new List<KeyValuePair<string, string>>();
        public KeyValuePair<string, string> ScghName;
        public LutItem LutItem;
        public List<PostProcessSheetRow> Cpp = new List<PostProcessSheetRow>();

        public PatternGenItem()
        {
            ScghName = new KeyValuePair<string, string>();
            InitDictionary = new List<string>();
            Patset = new KeyValuePair<string, string>();
            PatSubr = new List<KeyValuePair<string, string>>();
        }
    }
}
