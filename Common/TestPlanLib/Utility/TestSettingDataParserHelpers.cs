using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

using CommonLib.Extension;

using TestPlanLib.Const;
using TestPlanLib.DataStruct;
using TestPlanLib.HardIpDc.BaseData;

namespace TestPlanLib.Utility
{
    public static class TestSettingDataParserHelpers
    {
        public static string GetPowerMergeTableName(string job)
        {
            if (job.ContainsIgnoreCase("ft"))
            {
                return "FT";
            }

            return job.ContainsIgnoreCase("cp") ? "CP" : "Others";
        }

        public static void AddHardipDcCategory(HardIpDcSheet hardIpDcSheet, TestSettingData testSettingData)
        {
            List<DcCategoryName> dcCategorys = testSettingData.DcCategorys;
            List<TestSettingRow> rows = testSettingData.DataRows;

            foreach (HardIpCategoryDef hardipCategoryDef in hardIpDcSheet.Rows)
            {
                //Add Category Header
                string hardipDcCategory = "HardIP_" + DcCategoryName.CategoryDefaultValue + "_" + DcCategoryName.CategoryDefaultValue + "_" + DcCategoryName.CategoryDefaultValue + "_" + hardipCategoryDef.CategoryName;
                if (dcCategorys.Exists(s => s.CategoryName.EqualsIgnoreCase(hardipDcCategory)) || string.IsNullOrEmpty(hardipCategoryDef.DcCategory))
                {
                    continue;
                }

                var hardipDcCategoryName = new DcCategoryName(hardipDcCategory, true)
                {
                    ColumnIndex = dcCategorys[^1].ColumnIndex + 1
                };
                dcCategorys.Add(hardipDcCategoryName);

                //Add Category Value
                foreach (TestSettingRow testSettingRow in rows)
                {
                    bool isValtRow = testSettingRow.PowerPinName.EndsWithIgnoreCase(TestPlanConst.ValtRowPinNameFlag);
                    string pinName = testSettingRow.PowerPinName.Replace(TestPlanConst.ValtRowPinNameFlag, "");

                    HardIpDcRow? hardipDcRow = hardipCategoryDef.DataRows.Find(s => s.PinName.EqualsIgnoreCase(pinName));
                    //if hardipDcRow is null or hardipDc value is null, use hardip default value

                    DcCategoryValue hardipDcCategoryValue = new DcCategoryValue(hardipDcCategoryName.CategoryName)
                    {
                        ColumnIndex = hardipDcCategoryName.ColumnIndex
                    };
                    DcCategoryItem nv = new DcCategoryItem(EnumCategorySoruceType.HardIP);
                    DcCategoryItem hv = new DcCategoryItem(EnumCategorySoruceType.HardIP);
                    DcCategoryItem lv = new DcCategoryItem(EnumCategorySoruceType.HardIP);
                    hardipDcCategoryValue.Nv = nv;
                    hardipDcCategoryValue.Lv = lv;
                    hardipDcCategoryValue.Hv = hv;
                    testSettingRow.DcCategoryValues.Add(hardipDcCategoryValue);

                    //NV Value
                    string hardipDcNv = "";
                    if (hardipDcRow != null)
                    {
                        hardipDcNv = isValtRow ? hardipDcRow.NvValt : hardipDcRow.Nv;
                    }

                    if (hardipDcRow != null && !string.IsNullOrEmpty(hardipDcNv))
                    {
                        nv.Value = hardipDcNv;
                    }

                    //LV Value
                    if (hardipDcRow != null && !string.IsNullOrEmpty(hardipDcRow.LvRatio))
                    {
                        string lvRation = ConvertRationToPercentage(hardipDcRow.LvRatio);
                        lv.Value = lvRation;
                    }

                    //HV Value
                    if (hardipDcRow != null && !string.IsNullOrEmpty(hardipDcRow.HvRatio))
                    {
                        string hvRation = ConvertRationToPercentage(hardipDcRow.HvRatio);
                        hv.Value = hvRation;
                    }
                }
            }
        }

        private static string ConvertRationToPercentage(string input)
        {
            if (!double.TryParse(input, out double dvalue))
            {
                throw new Exception("HV value in HardIpDc sheet format error:" + input);
            }

            dvalue = (dvalue - 1) * 100;
            return dvalue + "%";
        }

        public static void AddIoLevelDcCategory(IoLevelsSheet ioLevelsSheet, TestSettingData testSettingData)
        {
            List<TestSettingRow> rows = testSettingData.DataRows;
            List<DcCategoryName> dcCategorys = testSettingData.DcCategorys;
            DcCategoryName allCategoryName = FindLcdAllCategoryName(testSettingData) ?? throw new Exception("Can not find Lcd ALL categoryName in TestSetting!");

            List<string> ioLevelCategories = [.. ioLevelsSheet.RowList.SelectMany(p => p.IoLevelDate).Select(p => p.Level).Distinct()];
            foreach (string ioLevelCategory in ioLevelCategories)
            {
                //Add Category Header              
                if (dcCategorys.Exists(s => s.CategoryName.EqualsIgnoreCase(ioLevelCategory)))
                {
                    continue;
                }

                var ioLevelDcCategoryName = new DcCategoryName(ioLevelCategory)
                {
                    ColumnIndex = dcCategorys[^1].ColumnIndex + 1
                };
                dcCategorys.Add(ioLevelDcCategoryName);

                //Add Category Value
                foreach (TestSettingRow testSettingRow in rows)
                {
                    //use All_x_x_x value
                    DcCategoryValue allCategoryValue = testSettingRow.DcCategoryValues.Find(s => s.ColumnIndex == allCategoryName.ColumnIndex)!;

                    DcCategoryValue ioLevelDcCategoryValue = new DcCategoryValue(ioLevelDcCategoryName.CategoryName)
                    {
                        ColumnIndex = ioLevelDcCategoryName.ColumnIndex
                    };
                    DcCategoryItem nv = new DcCategoryItem(EnumCategorySoruceType.IOLevel);
                    DcCategoryItem hv = new DcCategoryItem(EnumCategorySoruceType.IOLevel);
                    DcCategoryItem lv = new DcCategoryItem(EnumCategorySoruceType.IOLevel);
                    ioLevelDcCategoryValue.Nv = nv;
                    ioLevelDcCategoryValue.Lv = lv;
                    ioLevelDcCategoryValue.Hv = hv;
                    testSettingRow.DcCategoryValues.Add(ioLevelDcCategoryValue);

                    nv.Value = allCategoryValue.Nv.Value;
                    lv.Value = allCategoryValue.Lv.Value;
                    hv.Value = allCategoryValue.Hv.Value;
                }
            }
        }

        private static DcCategoryName? FindLcdAllCategoryName(TestSettingData testSettingData)
        {
            List<DcCategoryName> dcCategorys = testSettingData.DcCategorys;
            DcCategoryName? hardipDefaultCategory = dcCategorys.Find(s => s.DcCategoryInfo.Test.EqualsIgnoreCase("All") && s.DcCategoryInfo.Domain.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) && s.DcCategoryInfo.Subtest.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) && s.DcCategoryInfo.PmodePatternVdip.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) && s.DcCategoryInfo.UserDefined.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue));
            return hardipDefaultCategory;
        }

        public static void FillDcCategoryByBase(DcCategoryValue dcCategoryValue)
        {
            DcCategoryItem sourceItem;
            DcCategoryItem nv = dcCategoryValue.Nv;
            DcCategoryItem hv = dcCategoryValue.Hv;
            DcCategoryItem lv = dcCategoryValue.Lv;

            if (string.IsNullOrEmpty(nv.OriginValue))
            {
                List<PropertyInfo> dcItems = [.. dcCategoryValue.GetType().GetProperties().Where(x => x.PropertyType == typeof(DcCategoryItem))];
                PropertyInfo? item = dcItems.Find(x =>
                {
                    DcCategoryItem currentItem = (DcCategoryItem)x.GetValue(dcCategoryValue, null)!;
                    string value = (string)currentItem.GetType().GetProperty("Value")!.GetValue(currentItem, null)!;
                    return !string.IsNullOrEmpty(value);
                });

                if (item == null)
                {
                    sourceItem = new DcCategoryItem
                    {
                        OriginValue = "",
                        Formula = ""
                    };
                }
                else
                {
                    sourceItem = (DcCategoryItem)item.GetValue(dcCategoryValue, null)!;
                }

                FillDcCategoryNv(sourceItem, nv);
            }
            else
            {
                sourceItem = nv;
            }

            if (string.IsNullOrEmpty(hv.OriginValue))
            {
                FillDcCategoryHv(sourceItem, hv);
            }

            if (string.IsNullOrEmpty(lv.OriginValue))
            {
                FillDcCategoryLv(sourceItem, lv);
            }
        }

        public static void FillDcCategoryByZero(DcCategoryValue dcCategoryValue)
        {
            DcCategoryItem? currentItem = dcCategoryValue.Nv;
            if (currentItem != null && string.IsNullOrEmpty(currentItem.OriginValue))
            {
                string hv = dcCategoryValue.Hv.OriginValue;
                string lv = dcCategoryValue.Lv.OriginValue;
                if (!string.IsNullOrEmpty(hv) && !string.IsNullOrEmpty(lv))
                {
                    if (!hv.Contains('%') && !lv.Contains('%'))
                    {
                        bool isHvNumeric = double.TryParse(hv, out double hvValue);
                        bool isLvNumeric = double.TryParse(lv, out double lvValue);
                        if (isHvNumeric && isLvNumeric)
                        {
                            double nvValue = (hvValue + lvValue) / 2;
                            currentItem.IsOriginal = false;
                            currentItem.FillDataSource = "(" + hvValue.ToString("G15", CultureInfo.InvariantCulture) + "+" + lvValue.ToString("G15", CultureInfo.InvariantCulture) + ")/2";
                            currentItem.SourceType = EnumCategorySoruceType.Half_HV_LV;
                            currentItem.Value = nvValue.ToString("G15", CultureInfo.InvariantCulture);
                        }
                    }
                }
                else
                {
                    currentItem.IsOriginal = false;
                    currentItem.FillDataSource = "0";
                    currentItem.SourceType = EnumCategorySoruceType.FilledWithZero;
                    currentItem.Value = "0";
                }
            }

            DcCategoryItem currentItemHv = dcCategoryValue.Hv;
            if (string.IsNullOrEmpty(currentItemHv.OriginValue))
            {
                currentItemHv.IsOriginal = false;
                currentItemHv.FillDataSource = "0";
                currentItemHv.SourceType = EnumCategorySoruceType.FilledWithZero;
                currentItemHv.Value = "0";
            }

            DcCategoryItem currentItemLv = dcCategoryValue.Lv;
            if (string.IsNullOrEmpty(currentItemLv.OriginValue))
            {
                currentItemLv.IsOriginal = false;
                currentItemLv.FillDataSource = "0";
                currentItemLv.SourceType = EnumCategorySoruceType.FilledWithZero;
                currentItemLv.Value = "0";
            }
        }

        private static void FillDcCategoryNv(DcCategoryItem sourceItem, DcCategoryItem nv)
        {
            nv.Value = sourceItem.OriginValue;
            nv.Formula = sourceItem.Formula;
            nv.IsOriginal = false;
            nv.FillDataSource = sourceItem.OriginValue;
            nv.SourceType = EnumCategorySoruceType.FilledWithBase;
        }

        private static void FillDcCategoryLv(DcCategoryItem sourceItem, DcCategoryItem lv)
        {
            lv.Value = sourceItem.OriginValue;
            lv.Formula = sourceItem.Formula;
            lv.IsOriginal = false;
            lv.FillDataSource = sourceItem.OriginValue;
            lv.SourceType = EnumCategorySoruceType.FilledWithBase;
        }

        private static void FillDcCategoryHv(DcCategoryItem sourceItem, DcCategoryItem hv)
        {
            hv.Value = sourceItem.OriginValue;
            hv.Formula = sourceItem.Formula;
            hv.IsOriginal = false;
            hv.FillDataSource = sourceItem.OriginValue;
            hv.SourceType = EnumCategorySoruceType.FilledWithBase;
        }

        public static string ConvertMiliValt(string value)
        {
            bool isNumeric = double.TryParse(value, out double result);
            if (isNumeric)
            {
                result /= 1000;
                return result.ToString("G15", CultureInfo.InvariantCulture);
            }
            else
            {
                return value;
            }
        }

        public static string ConvertValt(string value)
        {
            return value;
        }
    }
}
