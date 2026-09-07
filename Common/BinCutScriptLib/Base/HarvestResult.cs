using System.Collections.Generic;

using CommonLib.Extension;

namespace BinCutScriptLib.Base
{
    public class HarvestResult
    {
        public int Site;
        public Dictionary<string, string> HarvesFlags = new(StringExtensions.IgnoreCase);
        //For debug
        public Line.BinCutLineBase Line { get; set; } = new();
    }
}
