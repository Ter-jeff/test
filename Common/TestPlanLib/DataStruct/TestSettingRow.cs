using System.Collections.Generic;
using System.Linq;

namespace TestPlanLib.DataStruct
{
    public class TestSettingRow
    {
        public string PowerPinName { set; get; } = string.Empty;
        public List<DcCategoryValue> DcCategoryValues { set; get; } = [];

        public TestSettingRow() { }

        public TestSettingRow(TestSettingRow testSettingRow)
        {
            if (testSettingRow == null)
            {
                return;
            }

            PowerPinName = testSettingRow.PowerPinName;
            DcCategoryValues = [.. testSettingRow.DcCategoryValues.Select(x => new DcCategoryValue(x))];
        }

        public TestSettingRow Copy()
        {
            return new TestSettingRow(this);
        }
    }
}
