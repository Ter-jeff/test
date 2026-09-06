using System.Collections.Generic;

namespace IgxlLib.IgxlBase
{
    public class PSet : IgxlRow
    {
        public string Name = string.Empty;
        public string Pin = string.Empty;
        public string InstrumentType = string.Empty;
        public string ThisLargeHeading = string.Empty;
        public Dictionary<string, string> Parameters = [];
        public string Comment = string.Empty;
    }
}
