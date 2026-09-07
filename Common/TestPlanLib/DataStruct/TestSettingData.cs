using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using TestPlanLib.Const;

namespace TestPlanLib.DataStruct
{
    public class TestSettingData
    {
        private List<string> _powerPins = [];
        private List<string> _valtPowerPins = [];
        private List<string> _dcCategoryNames = [];
        private List<TestSettingRow> _powerPinRows = [];
        private List<TestSettingRow>? _vrsPinRows;

        public EnumTestSettingBasicUnit BasicUnit { set; get; } = EnumTestSettingBasicUnit.mV;
        public string Job { set; get; } = "";
        public string OriginJob { set; get; } = "";
        public string SheetName { set; get; } = "";
        public List<DcCategoryName> DcCategorys { set; get; } = [];
        public List<TestSettingRow> DataRows { set; get; } = [];
        public List<TestSettingFooterRow> Footer { set; get; } = [];
        public string TestSettingVersion { get; set; } = string.Empty;

        public TestSettingData() { }

        public TestSettingData(TestSettingData testSettingData)
        {
            if (testSettingData == null)
            {
                return;
            }

            BasicUnit = testSettingData.BasicUnit;
            Job = testSettingData.Job;
            OriginJob = testSettingData.OriginJob;
            SheetName = testSettingData.SheetName;
            DcCategorys = [.. testSettingData.DcCategorys.Select(x => x.Copy())];
            DataRows = [.. testSettingData.DataRows.Select(x => x.Copy())];
            Footer = [.. testSettingData.Footer.Select(x => x.Copy())];
            TestSettingVersion = testSettingData.TestSettingVersion;
        }

        public TestSettingData Copy()
        {
            return new TestSettingData(this);
        }

        public List<string> GetDcCategorys()
        {
            _dcCategoryNames = [];
            List<string> categoryNamesUpper = [.. DcCategorys.Select(x => x.CategoryName.ToUpper()).Distinct()];
            foreach (string cateogryNameUpper in categoryNamesUpper)
            {
                _dcCategoryNames.Add(DcCategorys.Find(x => x.CategoryName.EqualsIgnoreCase(cateogryNameUpper))!.CategoryName);
            }
            return _dcCategoryNames;
        }

        public List<string> GetPowerPins()
        {
            _powerPins = [];
            List<string> pins = [.. DataRows.Select(x => x.PowerPinName).Distinct()];

            pins.ForEach(pin =>
            {
                if (!pin.StartsWithIgnoreCase("Pins_") && !pin.EndsWithIgnoreCase("_Valt"))
                {
                    _powerPins.Add(pin);
                }
            });
            return _powerPins;
        }

        public List<string> GetValtPowerPins()
        {
            _valtPowerPins = [];
            List<string> valtPins = GetValtPinRows().ConvertAll(x => x.PowerPinName);
            valtPins.ForEach(pin => _valtPowerPins.Add(pin.ToUpper().Replace(TestPlanConst.ValtRowPinNameFlag.ToUpper(), "").Trim()));
            return _valtPowerPins;
        }

        public List<TestSettingRow> GetPowerPinRows()
        {

            _powerPinRows = [.. DataRows.Where(x => !x.PowerPinName.StartsWithIgnoreCase("Pins_") && !x.PowerPinName.EndsWithIgnoreCase(TestPlanConst.ValtRowPinNameFlag))];

            return _powerPinRows;
        }

        public List<TestSettingRow> GetValtPinRows()
        {
            return _vrsPinRows ??= [.. DataRows.Where(x => x.PowerPinName.EndsWithIgnoreCase(TestPlanConst.ValtRowPinNameFlag))];
        }

        public static bool GetValuebyTestSettingRow(string categoryName, CategoryValueType categoryValueType, TestSettingRow testSettingRow, out string value)
        {
            value = "";
            if (testSettingRow.DcCategoryValues.Any(x => x.CategoryName.EqualsIgnoreCase(categoryName)))
            {
                DcCategoryValue category = testSettingRow.DcCategoryValues.Find(x => x.CategoryName.EqualsIgnoreCase(categoryName))!;
                if (categoryValueType == CategoryValueType.NV)
                {
                    value = category.Nv.Value;
                    return true;
                }
                if (categoryValueType == CategoryValueType.HV)
                {
                    value = category.Hv.Value;
                    return true;
                }
                if (categoryValueType == CategoryValueType.LV)
                {
                    value = category.Lv.Value;
                    return true;
                }
            }
            return false;
        }

        public List<TestSettingRow> GetAllCategory(List<string> powerPinNameList)
        {
            var result = new List<TestSettingRow>();
            foreach (string powerPinName in powerPinNameList)
            {
                if (DataRows.Any(x => x.PowerPinName.EqualsIgnoreCase(powerPinName)))
                {
                    result.Add(DataRows.Find(x => x.PowerPinName.EqualsIgnoreCase(powerPinName))!);
                }
            }
            return result;
        }

        public string? GetPinValue(string pin, string categoryName, CategoryValueType categoryValueType)
        {
            if (DcCategorys.Exists(x => x.CategoryName.EqualsIgnoreCase(categoryName)
                && x.ValueType == categoryValueType))
            {
                TestSettingRow? row = DataRows.Find(x => x.PowerPinName.EqualsIgnoreCase(pin));
                if (row == null)
                {
                    return "";
                }

                DcCategoryValue? dcCategory = row.DcCategoryValues.Find(y => y.CategoryName.EqualsIgnoreCase(categoryName));
                if (dcCategory == null)
                {
                    return "";
                }

                string value = "";
                if (categoryValueType == CategoryValueType.NV)
                {
                    value = dcCategory.Nv.Value;
                }
                else if (categoryValueType == CategoryValueType.LV)
                {
                    value = dcCategory.Lv.Value;
                }
                else if (categoryValueType == CategoryValueType.HV)
                {
                    value = dcCategory.Hv.Value;
                }

                if (!string.IsNullOrEmpty(value))
                {
                    //Transfer voltage from mV to V for TestSetting's voltage
                    return (double.Parse(value) / 1000).ToString();
                }
                return "";
            }
            return null;
        }
    }
}
