using System.IO;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace CommonReaderLib.PatternListCsv
{
    public class PatternListRow : MyRow
    {
        public PatternListRow(string sheetName)
        {
            SheetName = sheetName;
        }

        public string Number { set; get; }
        public string Pattern { set; get; }
        public string UseNotUse { set; get; }
        public string TimeSetLatest { set; get; }
        public string FileVersions { set; get; }
        public string ActualTimeSetVersion { get; internal set; }

        public string TimeSet
        {
            get
            {
                if (TimeSetLatest.EqualsIgnoreCase("N/A") || TimeSetLatest.EqualsIgnoreCase("NA"))
                {
                    return "";
                }
                return Path.GetFileName(TimeSetLatest);
            }
        }

        public string PatternVersion
        {
            get
            {
                if (FileVersions.EqualsIgnoreCase("N/A") || FileVersions.EqualsIgnoreCase("NA"))
                {
                    return "";
                }
                string patternVersion = Path.GetFileName(Regex.Replace(Regex.Replace(FileVersions, ".gz$", "",
                    RegexOptions.IgnoreCase), ".atp$", "", RegexOptions.IgnoreCase));
                patternVersion = Regex.Replace(patternVersion, ".Pat$", "", RegexOptions.IgnoreCase);
                return patternVersion;
            }
        }
    }
}
