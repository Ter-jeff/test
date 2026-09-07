using System.Collections.Generic;

using CommonLib.Extension;

using IgxlLib.Const;
using IgxlLib.IgxlBase;

namespace IgxlLib.IgxlSheets
{
    public class AcSpecSheet : SpecSheetBase<AcSpec>
    {
        public override string SheetType => "DTACSpecSheet";
        public override string IgxlSheetName => IgxlSheetNames.AcSpec;

        public AcSpecSheet(string sheetName) : base(sheetName)
        {
        }

        public AcSpecSheet(string sheetName, List<string> categoryList, List<string> selectorNameList) : base(sheetName, categoryList, selectorNameList)
        {
        }

        public bool IsSymbolExist(string name)
        {
            foreach (AcSpec acSpecs in Rows)
            {
                if (acSpecs.Symbol.EqualsIgnoreCase(name))
                {
                    return true;
                }
            }

            return false;
        }

        public AcSpec AddAcSpecs(string symbol, string value, string typ, string min, string max)
        {
            //Write basic data
            var acSpecs = new AcSpec(symbol, GetSelectorList(), value);
            //Write Category
            foreach (string category in CategoryList)
            {
                var categoryInSpec = new CategoryInSpec(category, typ, min, max);
                acSpecs.AddCategory(categoryInSpec);
            }

            AddRow(acSpecs);
            return acSpecs;
        }

        private static List<Selector> GetSelectorList()
        {
            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("Typ", "Typ"));
            selectorList.Add(new Selector("Min", "Min"));
            selectorList.Add(new Selector("Max", "Max"));
            return selectorList;
        }
    }
}
