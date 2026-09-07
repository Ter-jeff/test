using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Base;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using LogLib.Utility;

using TestPlanLib.DataStruct;

namespace LcdLib.Basic
{
    [ExcludeFromCodeCoverage]
    public partial class DcSpecGeneratorLcd(MultiTestSettingSheetsSingleton multiTestSettingSheetsSingleton, IoLevelsSheet ioLevelsSheet, GlobalSpecSheet globalSpecSheet) : GeneratorBase
    {

        [GeneratedRegex("All_X_X", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        #region Field
        //private Regex _efusePinReg;
        private readonly MultiTestSettingSheetsSingleton _testSettings = multiTestSettingSheetsSingleton;
        private readonly IoLevelsSheet _ioLevelSheet = ioLevelsSheet;
        private readonly GlobalSpecSheet _glbSpecSheet = globalSpecSheet;

        #endregion
        #region Constructor
        #endregion

        #region Menber Function

        public Dictionary<string, List<DcSpec>> GetIoDcSpecs()
        {
            Dictionary<string, List<DcSpec>> result = [];
            foreach (TestSettingData entry in _testSettings.TestSettingSheetsList)
            {
                List<string> dcCategorys = entry.GetDcCategorys();
                //dcCategorys.AddRange(ioDcCategory);
                string key = entry.Job;
                result.Add(key, []);
                var ioPinGroups = _ioLevelSheet.RowList.Where(p => !string.IsNullOrEmpty(p.Domain)).Select(p => p.Domain).Distinct().ToList();

                foreach (string pinGrpName in ioPinGroups.Distinct())
                {
                    List<DcSpec> specs = CreateDcSpecWithIoLevels(pinGrpName, dcCategorys);

                    result[key].AddRange(specs);
                }
            }

            return result;
        }

        public Dictionary<string, List<DcSpec>> GetPowerDcSpecs()
        {
            Dictionary<string, List<DcSpec>> result = [];
            foreach (TestSettingData entry in _testSettings.TestSettingSheetsList)
            {
                List<TestSettingRow> powerPinRows = entry.GetPowerPinRows();
                List<string> dcCategorys = entry.GetDcCategorys();
                List<DcCategoryName> titles = entry.DcCategorys;

                string key = entry.Job;
                result.Add(key, []);
                foreach (TestSettingRow row in powerPinRows)
                {
                    List<DcSpec> specs = CreateClassifiedPowerDcSpec(row.PowerPinName, dcCategorys, row, titles);
                    result[key].AddRange(specs);
                }
                result[key][0].SpecialComment = entry.TestSettingVersion;
            }

            return result;
        }

        private List<DcSpec> CreateClassifiedPowerDcSpec(string pinName, List<string> testSettingDcCategorys, TestSettingRow testSettingRow, List<DcCategoryName> dcCategoryNames)
        {
            List<DcSpec> result = [];

            DcSpec baseSpec = new DcSpec(JoinPinComponent(PinNamePrefix, pinName, "VAR")) { SelectorList = GetSelectorList() };

            DcCategoryValue? dcValue = null;
            foreach (string categoryName in testSettingDcCategorys)
            {
                int columnIndex = dcCategoryNames.Find(x => x.CategoryName.EqualsIgnoreCase(categoryName))!.ColumnIndex;
                dcValue = testSettingRow.DcCategoryValues.Find(x => x.ColumnIndex == columnIndex);

                CategoryInSpec categoryItem = GetCategoryItem(dcValue!, categoryName);
                baseSpec.AddCategory(categoryItem);
            }

            IEnumerable<string> extraCategories = _ioLevelSheet.RowList.SelectMany(p => p.IoLevelDate).Select(p => p.Level).Distinct();
            CategoryInSpec refDcCate = baseSpec.CategoryList.First(p => MyRegex().IsMatch(p.Name));
            foreach (string categoryName in extraCategories)
            {
                if (baseSpec.CategoryList.FindIndex(p => p.Name.EqualsIgnoreCase(categoryName)) != -1)
                {
                    continue;
                }

                CategoryInSpec categoryItem = new CategoryInSpec(categoryName, refDcCate.Typ!, refDcCate.Min!, refDcCate.Max!);

                baseSpec.AddCategory(categoryItem);
            }

            result.Add(baseSpec);

            return result;
        }

        private static CategoryInSpec GetCategoryItem(DcCategoryValue dcCategoryValue, string categoryName)
        {
            string typ = dcCategoryValue.Nv.Value.Contains('%') ? PercentageConvertToString(dcCategoryValue.Nv.Value) : dcCategoryValue.Nv.Value;
            string max = dcCategoryValue.Hv.Value.Contains('%') ? MakeCeilingFormula(PercentageConvertToString(dcCategoryValue.Hv.Value, typ)) : dcCategoryValue.Hv.Value;
            string min = dcCategoryValue.Lv.Value.Contains('%') ? MakeFloorFormula(PercentageConvertToString(dcCategoryValue.Lv.Value, typ)) : dcCategoryValue.Lv.Value;

            CategoryInSpec categoryItem = new CategoryInSpec(categoryName, FormulaPrefix + typ, FormulaPrefix + min, FormulaPrefix + max);
            return categoryItem;
        }

        private static string PercentageConvertToString(string formula, string nv = "1")
        {
            string answer;
            try
            {
                double percent = (double.Parse(formula.TrimEnd('%')) * 0.01) + 1;
                answer = nv + FormulaConnector + percent.ToString("G15", CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                ErrorMessageBox.Show(string.Format(ex.ToString()));
                answer = "error";
            }
            return answer;
        }

        private List<DcSpec> CreateDcSpecWithIoLevels(string pinName, List<string> dcCategorys)
        {
            List<DcSpec> result = [];
            DcSpec specIo = new DcSpec(pinName + "_VAR");
            foreach (string category in dcCategorys)
            {
                CategoryInSpec categroyItem = new CategoryInSpec(category);
                string glbSymbol = SpecFormat.GenGlbSpecSymbol(pinName);
                List<GlobalSpec> row = _glbSpecSheet.FindRowList(glbSymbol, "");
                if (row.Count == 0)
                {
                    glbSymbol = SpecFormat.GenGlbSpecSymbol(pinName);
                }

                categroyItem.Typ = "=_" + glbSymbol;
                categroyItem.Max = SpecFormat.GenGlbRatio(glbSymbol, "IO_Pins_GLB_Plus");
                categroyItem.Min = SpecFormat.GenGlbRatio(glbSymbol, "IO_Pins_GLB_Minus");
                specIo.AddCategory(categroyItem);
            }
            specIo.SelectorList = GetSelectorList();
            result.Add(specIo);

            return result;
        }

        private static List<Selector> GetSelectorList()
        {
            List<Selector> result =
            [
                new Selector("Min", "Min"),
                new Selector("Typ", "Typ"),
                new Selector("Max", "Max")
            ];
            return result;
        }

        private static string MakeCeilingFormula(string pCell1)
        {
            //=$B$5
            string lStrResult;
            _ = double.TryParse(LocalSpecs.PwrSupplyRes, out double outD);
            if (outD >= 0.001)
            {
                lStrResult = "CEILING(" + pCell1 + "," + outD.ToString("G15", CultureInfo.InvariantCulture) + ")";
            }
            else
            {
                lStrResult = pCell1;
            }

            return lStrResult;
        }

        private static string MakeFloorFormula(string pCell1)
        {
            _ = double.TryParse(LocalSpecs.PwrSupplyRes, out double outD);
            //=$B$5
            string lStrResult;
            if (outD >= 0.001)
            {
                lStrResult = "FLOOR(" + pCell1 + "," + outD.ToString("G15", CultureInfo.InvariantCulture) + ")";
            }
            else
            {
                lStrResult = pCell1;
            }

            return lStrResult;
        }
        #endregion
    }
}
