namespace TestPlanLib.DataStruct
{
    public class DcCategoryName(string categoryName, bool isHardIpDcCategpry = false)
    {
        #region filed
        /// <summary>
        /// "X"
        /// </summary>
        public const string CategoryDefaultValue = "X";

        #endregion
        #region constructor

        #endregion

        #region property

        public DcCategoryInfo DcCategoryInfo { get; } = new DcCategoryInfo(categoryName, isHardIpDcCategpry);

        public string CategoryName { get; } = categoryName;

        public int ColumnIndex { set; get; }

        public CategoryValueType ValueType { set; get; }

        #endregion

        public DcCategoryName(DcCategoryName dcCategoryName) : this(dcCategoryName?.CategoryName ?? "", dcCategoryName?.DcCategoryInfo.IsHardipDcCategory ?? false)
        {
            if (dcCategoryName == null)
            {
                return;
            }

            ColumnIndex = dcCategoryName.ColumnIndex;
            ValueType = dcCategoryName.ValueType;
        }

        public DcCategoryName Copy()
        {
            return new DcCategoryName(this);
        }

        public string GetvoltageType()
        {
            if (ValueType == CategoryValueType.HV)
            {
                return "Max";
            }
            if (ValueType == CategoryValueType.LV)
            {
                return "Min";
            }
            return ValueType == CategoryValueType.NV ? "Typ" : "";
        }
    }
}
