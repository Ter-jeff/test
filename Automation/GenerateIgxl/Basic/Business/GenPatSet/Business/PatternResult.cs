using System.Collections.Generic;

namespace Automation.GenerateIgxl.Basic.Business.GenPatSet.Business
{
    public class PatternResult
    {
        public string FilePath { get; set; }
        public string Pattern { get; set; }
        public string TimeSet { get; set; } //other
        public bool HasPat { get; set; } //PAT
        public string VmVectorName { get; set; } //PAT
        public string OpcodeMode { get; set; } //PAT
        public int PatternVectorCount; //PAT
        public int NumOfSvm { get; set; } //PAT
        public int NumOfLvm { get; set; } //PAT
        public int LicenseCount { get; set; } //PAT

        public List<string> PinList = new List<string>(); //PAT
        public List<string> ModuleNameList = new List<string>(); //PAT
        public Dictionary<string, string> ScanInList = new Dictionary<string, string>(); //PAT

    }
}
