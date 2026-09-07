namespace TestPlanLib.DataStruct
{
    public class DcCategoryItem
    {
        public string OriginValue { set; get; } = "";
        public string Value { set; get; } = "";
        public string Formula { set; get; } = "";

        public bool IsOriginal { get; set; } = true;
        public string FillDataSource { get; set; } = string.Empty;
        public EnumCategorySoruceType SourceType { get; set; }
        public EnumCategoryDisplayType DisplayType { get; set; }

        public DcCategoryItem(EnumCategorySoruceType enumCategorySoruceType = EnumCategorySoruceType.TestPlan, EnumCategoryDisplayType enumCategoryDisplayType = EnumCategoryDisplayType.Origin)
        {
            SourceType = enumCategorySoruceType;
            DisplayType = enumCategoryDisplayType;
        }

        public DcCategoryItem(DcCategoryItem dcCategoryItem)
        {
            if (dcCategoryItem == null)
            {
                return;
            }

            OriginValue = dcCategoryItem.OriginValue;
            Value = dcCategoryItem.Value;
            Formula = dcCategoryItem.Formula;
            IsOriginal = dcCategoryItem.IsOriginal;
            FillDataSource = dcCategoryItem.FillDataSource;
            SourceType = dcCategoryItem.SourceType;
            DisplayType = dcCategoryItem.DisplayType;
        }

        public DcCategoryItem Copy()
        {
            return new DcCategoryItem(this);
        }
    }
}
