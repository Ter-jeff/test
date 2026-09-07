using System.Collections.Generic;

namespace IgxlLib.IgxlBase
{
    public abstract class Spec : IgxlRow
    {
        public string Symbol { get; set; }
        public string Value { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public List<CategoryInSpec> CategoryList { get; protected set; } = [];
        public List<Selector> SelectorList { get; set; } = [];

        protected Spec(string symbol)
        {
            Symbol = symbol;
        }

        protected Spec(string symbol, string value = "", string comment = "")
        {
            Symbol = symbol;
            Value = string.IsNullOrEmpty(value) ? "0" : value;
            Comment = comment;
        }
    }
}
