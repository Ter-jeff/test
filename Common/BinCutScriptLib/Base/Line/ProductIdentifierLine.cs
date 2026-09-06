using System;

namespace BinCutScriptLib.Base.Line
{
    public class ProductIdentifierLine : BinCutLineBase
    {
        public ProductIdentifierLineRow GetIdentifierLineRow()
        {
            var productIdentifierRow = new ProductIdentifierLineRow { Line = this };
            string[] spilts = Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < spilts.Length; i++)
            {
                if (spilts[i].Contains("SITE", StringComparison.OrdinalIgnoreCase))
                {
                    string[] s = spilts[i].Split(['(', ')'], StringSplitOptions.RemoveEmptyEntries);

                    _ = int.TryParse(s[1], out productIdentifierRow.Site);
                    continue;
                }
                if (spilts[i] == "Product_Identifier")
                {
                    _ = int.TryParse(spilts[i + 2], out productIdentifierRow.ProductIdentifier);
                }
                if (productIdentifierRow.Site != -1 && productIdentifierRow.ProductIdentifier != -1)
                {
                    break;
                }
            }
            return productIdentifierRow;
        }
    }

    public class ProductIdentifierLineRow
    {
        public int Site = -1;
        public int ProductIdentifier = -1;
        public BinCutLineBase Line = new();
    }
}
