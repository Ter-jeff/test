using System.Collections.Generic;

namespace IgxlLib.IgxlBase
{
    public class AcSpec : Spec
    {
        public AcSpec(string acSpecSymbol, List<Selector>? selectors, string value = "", string comment = "") : base(acSpecSymbol, value, comment)
        {
            CategoryList = [];
            SelectorList = selectors ?? [];
        }

        public void AddCategory(CategoryInSpec categoryInSpec)
        {
            CategoryList.Add(categoryInSpec);
        }

        public bool ContainsCategory(string categoryName)
        {
            return CategoryList.Exists(p => p.Name == categoryName);
        }

        public CategoryInSpec? GetCategoryItem(string categoryName)
        {
            return CategoryList.Find(p => p.Name == categoryName);
        }
    }
}
