namespace TestPlanLib.DataStruct
{
    public class TestSettingFooterRow
    {
        public string PinName { set; get; } = "";
        public string Value { set; get; } = "";
        public string Formula { set; get; } = "";
        public string Comment { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;

        public TestSettingFooterRow() { }

        public TestSettingFooterRow(TestSettingFooterRow testSettingFooterRow)
        {
            if (testSettingFooterRow == null)
            {
                return;
            }

            PinName = testSettingFooterRow.PinName;
            Value = testSettingFooterRow.Value;
            Formula = testSettingFooterRow.Formula;
            Comment = testSettingFooterRow.Comment;
            Address = testSettingFooterRow.Address;
        }

        public TestSettingFooterRow Copy()
        {
            return new TestSettingFooterRow(this);
        }
    }
}
