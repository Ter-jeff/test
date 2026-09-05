using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

namespace TestPlanLib
{
    public class VreMbistLookupTable
    {
        public List<OreMbistLookupRow> Rows = new List<OreMbistLookupRow>();
        public Dictionary<string, int> HeaderIndex;
        public string SheetName;
        public bool HarvestJudgement(string patternName)
        {
            return Rows.Any(row => patternName.ContainsIgnoreCase(row.Server) && patternName.ContainsIgnoreCase(row.MemoryGroup) && (string.IsNullOrEmpty(row.ExcludePattern) || !patternName.Contains(row.ExcludePattern)) && patternName.Split('_')[9].StartsWith(row.Pmode));
        }
    }

    public class OreMbistLookupRow
    {
        public string Harvesting { get; set; }
        public string Server { get; set; }
        public string Pmode { get; set; }
        public string MemoryGroup { get; set; }
        public string ExcludePattern { get; set; }
        public List<string> ExcludePatterns => ExcludePattern.Split(';').ToList();

        public List<string> Servers => Server.Split(',').ToList();
        public List<string> MemoryGroups => MemoryGroup.Split(',').ToList();

    }
}
