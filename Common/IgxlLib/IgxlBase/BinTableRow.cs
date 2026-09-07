using System.Collections.Generic;
using System.Linq;

namespace IgxlLib.IgxlBase
{
    public class BinTableRow : IgxlRow
    {
        public const string ResultFailStop = "Fail-stop";

        public string Name { get; set; } = string.Empty;
        public string ItemList { get; set; } = string.Empty;
        public string Op { get; set; } = string.Empty;
        public string Sort { get; set; } = string.Empty;
        public string Bin { get; set; } = string.Empty;
        public Dictionary<string, string> ExtraBinDictionary { get; set; } = [];
        public string Result { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public List<string> Items { get; set; } = [];
        public Dictionary<int, string> ItemsWithIndex { get; set; } = [];

        public BinTableRow(string sheetName)
        {
            SheetName = sheetName;
        }

        public BinTableRow()
        {
        }

        public BinTableRow(BinTableRow binTableRow) : base(binTableRow)
        {
            if (binTableRow == null)
            {
                return;
            }

            Name = binTableRow.Name;
            ItemList = binTableRow.ItemList;
            Op = binTableRow.Op;
            Sort = binTableRow.Sort;
            Bin = binTableRow.Bin;
            Result = binTableRow.Result;
            Comment = binTableRow.Comment;

            ExtraBinDictionary = binTableRow.ExtraBinDictionary?.ToDictionary(k => k.Key, v => v.Value)
                                      ?? [];

            Items = binTableRow.Items?.ToList()
                         ?? [];

            ItemsWithIndex = binTableRow.ItemsWithIndex?.ToDictionary(k => k.Key, v => v.Value)
                                  ?? [];
        }

        public BinTableRow Copy()
        {
            return new BinTableRow(this);
        }

        public string GetUniqKey()
        {
            string items = string.Join(",", Items);
            return $"{Name.ToLower()};{ItemList.ToLower()};{items.ToLower()}";
        }
    }
}
