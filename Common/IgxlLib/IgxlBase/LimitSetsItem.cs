using System.Diagnostics;

namespace IgxlLib.IgxlBase
{
    [DebuggerDisplay("{CategoryName} ")]
    public class LimitSetsItem
    {
        public string CategoryName = string.Empty;
        public string LoLim = "NA";
        public string HiLim = "NA";

        public LimitSetsItem()
        {
        }

        public LimitSetsItem(string categoryName, string lo, string hi)
        {
            CategoryName = categoryName;
            if (!string.IsNullOrEmpty(lo))
            {
                LoLim = lo;
            }

            if (!string.IsNullOrEmpty(hi))
            {
                HiLim = hi;
            }
        }
    }
}
