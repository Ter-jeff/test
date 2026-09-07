using System.Collections.Generic;

namespace IgxlLib.IgxlBase
{
    public class DcSpec : AcSpec
    {
        public string SpecialComment { get; set; } = string.Empty;

        public int CategoryCount
        {
            get { return CategoryList.Count; }
        }

        public DcSpec(string dcSpecSymbol, string value = "", string comment = "") : base(dcSpecSymbol, null, value, comment)
        {
            CategoryList = [];
            SelectorList = [];
        }

        public DcSpec(string dcSpecSymbol, List<Selector> selectors, string value = "", string comment = "") : base(dcSpecSymbol, selectors, value, comment)
        {
            CategoryList = [];
            SelectorList = selectors;
        }
    }
}
