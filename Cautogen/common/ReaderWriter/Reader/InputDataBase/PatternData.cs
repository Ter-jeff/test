using System.IO;
using System.Text.RegularExpressions;

namespace Cautogen.common.ReaderWriter.Reader.InputDataBase
{
    public class PatternData
    {
        public string PatternName { get; set; }
        public string LatestVersion { get; set; }
        public string Use { get; set; }
        public string Org { get; set; }
        public string TypeSpec { get; set; }
        public string TimesetVersion { get; set; }
        public string FileVersion { get; set; }
        public string OpCode { get; set; }
        public string ScanMode { get; set; }
        public string Halt { get; set; }
        public string OriginalTimingMode { get; set; }
        public string Check { get; set; }
        public string TpCategory { get; set; }
        public bool IsExist { get; set; }

        private readonly Regex _rePatVersion = new Regex(@"_(?<patVersion>[0-9]+_[a-zA-Z0-9]+_[0-9]+)\.ATP", RegexOptions.IgnoreCase);

        public PatternData()
        {
            PatternName = "";
            LatestVersion = "";
            Use = "";
            Org = "";
            TypeSpec = "";
            TimesetVersion = "";
            FileVersion = "";
            OpCode = "";
            ScanMode = "";
            Halt = "";
            OriginalTimingMode = "";
            Check = "";
            TpCategory = "";
            IsExist = false;
        }

        public string GetPatVersion()
        {
            Match match = _rePatVersion.Match(FileVersion);
            return match.Success ? match.Groups["patVersion"].Value : "N/A";
        }

        public string GetFullName()
        {
            Match match = _rePatVersion.Match(FileVersion);
            if (match.Success)
            {
                return $"{PatternName}_{match.Groups["patVersion"].Value}";
            }
            else
            {
                return "N/A";
            }
        }

        public string GetFileVersionNameOnly()
        {
            return Path.GetFileName(FileVersion).ToUpper().Replace(".ATP.GZ", "").Replace(".PAT.GZ", "");
        }
    }
}
