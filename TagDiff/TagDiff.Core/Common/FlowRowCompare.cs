using IgxlLib.IgxlBase;

namespace TagDiff.Core.Common
{
    public class FlowRowCompare : FlowRow
    {
        public bool IsChecked { get; set; } = false;
        public string BaseSheet { get; set; } = string.Empty;
        public string BaseRow { get; set; } = string.Empty;
        public string ComparedSheet { get; set; } = string.Empty;
        public string ComparedRow { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
    }
}
