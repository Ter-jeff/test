namespace TestPlanLib.DataStruct
{
    public class DcCategoryValue
    {
        public DcCategoryItem Nv { set; get; } = new();
        public DcCategoryItem Hv { set; get; } = new();
        public DcCategoryItem Lv { set; get; } = new();
        public string CategoryName { get; } = "";
        public int ColumnIndex { get; set; }

        public DcCategoryValue(string categoryName)
        {
            Nv = new DcCategoryItem();
            Hv = new DcCategoryItem();
            Lv = new DcCategoryItem();
            CategoryName = categoryName;
        }

        public DcCategoryValue(DcCategoryValue dcCategoryValue)
        {
            if (dcCategoryValue == null)
            {
                return;
            }

            // Initialize read-only property
            CategoryName = dcCategoryValue.CategoryName;
            ColumnIndex = dcCategoryValue.ColumnIndex;

            // Deep copy the custom items
            Nv = dcCategoryValue.Nv.Copy();
            Hv = dcCategoryValue.Hv.Copy();
            Lv = dcCategoryValue.Lv.Copy();
        }

        public DcCategoryValue Copy()
        {
            return new DcCategoryValue(this);
        }
    }
}
