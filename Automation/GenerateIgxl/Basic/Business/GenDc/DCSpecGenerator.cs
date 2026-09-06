using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Base;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Utility;

using IgxlLib.IgxlBase;

using LogLib.Static;
using LogLib.Utility;

using TestPlanLib.DataStruct;

namespace Automation.GenerateIgxl.Basic.Business.GenDc
{
    public class DcSpecGenerator : GeneratorBase
    {
        private readonly MultiTestSettingSheetsSingleton _testSettings;
        private readonly IoInfoSheet _ioInfo;
        private readonly IoInfoSheet _ioInfoConcurrent;

        public static readonly Regex Reg = new Regex("[a-z]", RegexOptions.IgnoreCase);

        public DcSpecGenerator(MultiTestSettingSheetsSingleton testSettings, IoInfoSheet ioInfo, IoInfoSheet ioInfoConcurrent)
        {
            _testSettings = testSettings;
            _ioInfo = ioInfo;
            _ioInfoConcurrent = ioInfoConcurrent;
        }

        public Dictionary<string, List<DcSpec>> GetIoDcSpecsByBlock()
        {
            var result = new Dictionary<string, List<DcSpec>>();
            foreach (TestSettingData entry in _testSettings.TestSettingSheetsList)
            {
                var nonScanCategories = new HashSet<string>();
                var scanCategories = new HashSet<string>();
                foreach (DcCategoryName dcCategory in entry.DcCategorys)
                {
                    if (dcCategory.CategoryName.StartsWith("TD_", StringComparison.CurrentCultureIgnoreCase) ||
                        dcCategory.CategoryName.StartsWith("PDF_", StringComparison.CurrentCultureIgnoreCase) ||
                        dcCategory.CategoryName.StartsWith("SAChain_", StringComparison.CurrentCultureIgnoreCase) ||
                        dcCategory.CategoryName.StartsWith("SA_", StringComparison.CurrentCultureIgnoreCase))
                    {
                        scanCategories.Add(dcCategory.CategoryName);
                    }
                    else
                    {
                        nonScanCategories.Add(dcCategory.CategoryName);
                    }
                }

                string key = entry.Job;

                var specs = new List<DcSpec>();
                var specsScan = new List<DcSpec>();
                foreach (IoInfoRow row in _ioInfo.GetBlockIoInfo("Common"))
                {
                    specs.AddRange(CreateClassifiedIoDcSpec(row, "", nonScanCategories.ToList()));
                    if (!_ioInfo.GetBlockIoInfo("Scan").Exists(x => x.PinGrpName.Equals(row.PinGrpName)))
                    {
                        specsScan.AddRange(CreateClassifiedIoDcSpec(row, "SC", scanCategories.ToList()));
                    }
                }
                foreach (IoInfoRow row in _ioInfo.GetBlockIoInfo("Scan"))
                {
                    specsScan.AddRange(CreateClassifiedIoDcSpec(row, "SC", scanCategories.ToList()));
                }
                foreach (IoInfoRow row in _ioInfo.GetBlockIoInfo("DCTEST_Continuity"))
                {
                    specs.AddRange(CreateClassifiedIoDcSpec(row, "", nonScanCategories.ToList()));
                }

                if (_ioInfoConcurrent != null)
                {
                    foreach (IoInfoRow row in _ioInfoConcurrent.GetBlockIoInfo("Common"))
                    {
                        specs.AddRange(CreateClassifiedIoDcSpec(row, "", nonScanCategories.ToList()));
                        specsScan.AddRange(CreateClassifiedIoDcSpec(row, "SC", scanCategories.ToList()));
                    }
                    foreach (IoInfoRow row in _ioInfoConcurrent.GetBlockIoInfo("Scan"))
                    {
                        specsScan.AddRange(CreateClassifiedIoDcSpec(row, "SC", scanCategories.ToList(), "Scan"));
                    }
                    foreach (IoInfoRow row in _ioInfoConcurrent.GetBlockIoInfo("HardIP"))
                    {
                        specs.AddRange(CreateClassifiedIoDcSpec(row, "", nonScanCategories.ToList()));
                    }
                    foreach (IoInfoRow row in _ioInfoConcurrent.GetBlockIoInfo("DCTEST_Continuity"))
                    {
                        specs.AddRange(CreateClassifiedIoDcSpec(row, "", nonScanCategories.ToList()));
                    }
                }
                result.Add(key, specs);
                result.Add(key + "_SC", specsScan);
            }

            return result;
        }


        public Dictionary<string, List<DcSpec>> GetIoDcSpecs()
        {
            var result = new Dictionary<string, List<DcSpec>>();
            foreach (TestSettingData entry in _testSettings.TestSettingSheetsList)
            {
                List<string> list = entry.GetDcCategorys();
                string key = entry.Job;
                result.Add(key, new List<DcSpec>());
                var totalIoInfos = new List<IoInfoRow>();
                totalIoInfos.AddRange(_ioInfo.GetBlockIoInfo("Common"));
                totalIoInfos.AddRange(_ioInfo.GetBlockIoInfo("Scan").Where(x => !totalIoInfos.Exists(y => y.PinGrpName.Equals(x.PinGrpName))));
                totalIoInfos.AddRange(_ioInfo.GetBlockIoInfo("HardIP").Where(x => !totalIoInfos.Exists(y => y.PinGrpName.Equals(x.PinGrpName))));
                totalIoInfos.AddRange(_ioInfo.GetBlockIoInfo("DCTEST_Continuity").Where(x => !totalIoInfos.Exists(y => y.PinGrpName.Equals(x.PinGrpName))));

                if (_ioInfoConcurrent != null)
                {
                    totalIoInfos.AddRange(_ioInfoConcurrent.GetBlockIoInfo("Common"));
                    totalIoInfos.AddRange(_ioInfoConcurrent.GetBlockIoInfo("Scan").Where(x => !totalIoInfos.Exists(y => y.PinGrpName.Equals(x.PinGrpName))));
                    totalIoInfos.AddRange(_ioInfoConcurrent.GetBlockIoInfo("HardIP").Where(x => !totalIoInfos.Exists(y => y.PinGrpName.Equals(x.PinGrpName))));
                    totalIoInfos.AddRange(_ioInfoConcurrent.GetBlockIoInfo("DCTEST_Continuity").Where(x => !totalIoInfos.Exists(y => y.PinGrpName.Equals(x.PinGrpName))));
                }

                foreach (IoInfoRow row in totalIoInfos)
                {
                    List<DcSpec> specs = CreateClassifiedIoDcSpec(row, "", list);

                    result[key].AddRange(specs);
                }
            }

            return result;
        }

        public Dictionary<string, List<DcSpec>> GetPowerDcSpecsByBlock()
        {
            var result = new Dictionary<string, List<DcSpec>>();
            foreach (TestSettingData entry in _testSettings.TestSettingSheetsList)
            {
                List<TestSettingRow> powerPinRows = entry.GetPowerPinRows();
                List<TestSettingRow> valtPinRows = entry.GetValtPinRows();
                List<DcCategoryName> titles = entry.DcCategorys;
                string key = entry.Job;
                var hashSet = new HashSet<string>();
                var dcCategoryScan = new HashSet<string>();
                foreach (DcCategoryName dcCategory in entry.DcCategorys)
                {
                    if (dcCategory.CategoryName.StartsWith("TD_", StringComparison.CurrentCultureIgnoreCase) || dcCategory.CategoryName.StartsWith("PDF_", StringComparison.CurrentCultureIgnoreCase) || dcCategory.CategoryName.StartsWith("SAChain_", StringComparison.CurrentCultureIgnoreCase) ||
                        dcCategory.CategoryName.StartsWith("SA_", StringComparison.CurrentCultureIgnoreCase))
                    {
                        dcCategoryScan.Add(dcCategory.CategoryName);
                    }
                    else
                    {
                        hashSet.Add(dcCategory.CategoryName);
                    }
                }

                var specs = new List<DcSpec>();
                var specsScan = new List<DcSpec>();
                foreach (TestSettingRow row in powerPinRows)
                {
                    specs.AddRange(CreateClassifiedPowerDcSpec(row.PowerPinName, "", hashSet.ToList(), row, valtPinRows, titles));
                    specsScan.AddRange(CreateClassifiedPowerDcSpec(row.PowerPinName, "SC", dcCategoryScan.ToList(), row, valtPinRows, titles));
                }
                result.Add(key, specs);
                result.Add(key + "_SC", specsScan);
                result[key][0].SpecialComment = LocalSpecs.VoltageTbFileName.Find(x => x.Contains(key)) != null ? LocalSpecs.VoltageTbFileName.Find(x => x.Contains(key)).Split(new[] { '\\', '/' }).Last() : entry.TestSettingVersion;
                result[key + "_SC"][0].SpecialComment = LocalSpecs.VoltageTbFileName.Find(x => x.Contains(key)) != null ? LocalSpecs.VoltageTbFileName.Find(x => x.Contains(key)).Split(new[] { '\\', '/' }).Last() : entry.TestSettingVersion;
            }

            return result;
        }

        public Dictionary<string, List<DcSpec>> GetPowerDcSpecs()
        {
            var result = new Dictionary<string, List<DcSpec>>();
            foreach (TestSettingData entry in _testSettings.TestSettingSheetsList)
            {
                List<TestSettingRow> powerPinRows = entry.GetPowerPinRows();
                List<TestSettingRow> valtPinRows = entry.GetValtPinRows();
                List<string> dcCategory = entry.GetDcCategorys();
                List<DcCategoryName> titles = entry.DcCategorys;

                string key = entry.Job;
                result.Add(key, new List<DcSpec>());
                foreach (TestSettingRow row in powerPinRows)
                {
                    List<DcSpec> specsC = CreateClassifiedPowerDcSpec(row.PowerPinName, "", dcCategory, row, valtPinRows, titles);
                    result[key].AddRange(specsC);
                }
                result[key][0].SpecialComment = LocalSpecs.VoltageTbFileName.Find(x => x.Contains(key)) != null ? LocalSpecs.VoltageTbFileName.Find(x => x.Contains(key)).Split(new[] { '\\', '/' }).Last() : entry.TestSettingVersion;
            }
            return result;
        }

        private List<DcSpec> CreateClassifiedPowerDcSpec(string pinName, string suffix, List<string> list, TestSettingRow sourceData, List<TestSettingRow> vrsPinRows, List<DcCategoryName> titles, string chiplet = "")
        {
            var result = new List<DcSpec>();
            var baseSpec = new DcSpec(JoinPinComponent(PinNamePrefix, pinName, Combination.CombineByUnderLine(Combination.CombineByUnderLine("VAR", suffix), chiplet)))
            {
                SelectorList = GetSelectorList()
            };

            var vopSpec = new DcSpec(JoinPinComponent(PinNamePrefix, pinName, "VOP", Combination.CombineByUnderLine(Combination.CombineByUnderLine("VAR", suffix), chiplet)))
            {
                SelectorList = GetSelectorList()
            };

            TestSettingRow vrsRow = vrsPinRows.Find(x => x.PowerPinName.Equals(pinName + MultiTestSettingSheetsSingleton.ValtRowPinNameFlag, StringComparison.OrdinalIgnoreCase));

            foreach (string categoryName in list)
            {
                int columnIndex = titles.Find(x => x.CategoryName.Equals(categoryName, StringComparison.OrdinalIgnoreCase)).ColumnIndex;
                DcCategoryValue dcValue = sourceData.DcCategoryValues.Find(x => x.ColumnIndex == columnIndex);
                DcCategoryValue vrsDcValue = vrsRow?.DcCategoryValues.Find(x => x.ColumnIndex == columnIndex);
                dcValue = ConvertSpecialCategoryValue(dcValue, vrsDcValue);
                CategoryInSpec categoryItem = GetCategoryItem(dcValue, categoryName);
                if (vrsRow == null)
                {
                    baseSpec.AddCategory(categoryItem);
                }
                else
                {
                    if (string.IsNullOrEmpty(vrsDcValue.Nv.Value))
                    {
                        Response.Report("Valt Pin " + vrsRow.PowerPinName + " default value.", EnumMessageLevel.Error, 0);
                        baseSpec.AddCategory(categoryItem);
                        vopSpec.AddCategory(categoryItem);
                    }
                    else
                    {
                        CategoryInSpec categoryInSpec = GetCategoryItem(vrsDcValue, categoryName);
                        baseSpec.AddCategory(categoryItem);
                        vopSpec.AddCategory(categoryInSpec);
                    }
                }
            }

            result.Add(baseSpec);
            if (vrsRow != null)
            {
                result.Add(vopSpec);
            }
            return result;
        }

        private CategoryInSpec GetCategoryItem(DcCategoryValue dcValue, string categoryName)
        {
            string typ = dcValue.Nv.Value.Contains("%") ? PercentageConvertToString(dcValue.Nv.Value) : dcValue.Nv.Value;
            string max = dcValue.Hv.Value.Contains("%") ? MakeCeilingFormula(PercentageConvertToString(dcValue.Hv.Value, typ)) : dcValue.Hv.Value;
            string min = dcValue.Lv.Value.Contains("%") ? MakeFloorFormula(PercentageConvertToString(dcValue.Lv.Value, typ)) : dcValue.Lv.Value;
            var categoryItem = new CategoryInSpec(categoryName, FormulaPrefix + typ, FormulaPrefix + min, FormulaPrefix + max);
            return categoryItem;
        }

        internal DcCategoryValue ConvertSpecialCategoryValue(DcCategoryValue dcValue, DcCategoryValue vrsDcValue)
        {
            if (Reg.IsMatch(dcValue.Hv.Value) || Reg.IsMatch(dcValue.Lv.Value) || Reg.IsMatch(dcValue.Nv.Value))
            {
                DcCategoryValue newDcValue = dcValue.Copy();
                if (Reg.IsMatch(dcValue.Hv.Value))
                {
                    if (vrsDcValue == null)
                    {
                        newDcValue.Hv.Value = "0";
                    }
                    else
                    {
                        newDcValue.Hv.Value = vrsDcValue.Hv.Value;
                    }
                }
                if (Reg.IsMatch(dcValue.Lv.Value))
                {

                    if (vrsDcValue == null)
                    {
                        newDcValue.Lv.Value = "0";
                    }
                    else
                    {
                        newDcValue.Lv.Value = vrsDcValue.Lv.Value;
                    }
                }
                if (Reg.IsMatch(dcValue.Nv.Value))
                {
                    if (vrsDcValue == null)
                    {
                        newDcValue.Nv.Value = "0";
                    }
                    else
                    {
                        newDcValue.Nv.Value = vrsDcValue.Nv.Value;
                    }
                }
                return newDcValue;
            }

            return dcValue;
        }

        internal string PercentageConvertToString(string formula, string nv = "1")
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

        private List<DcSpec> CreateClassifiedIoDcSpec(IoInfoRow ioInfoRow, string suffix, List<string> list, string concurrentSuffix = "")
        {
            string pinName = ioInfoRow.PinGrpName;
            var result = new List<DcSpec>();
            string chipletFromGrpName = Regex.Match(pinName, @"\w+_(?<chiplet>[A-z]\d+$)").Groups["chiplet"].ToString();

            string suffixText = Combination.CombineByUnderLine("VAR", suffix);
            var specVih = new DcSpec(string.IsNullOrEmpty(concurrentSuffix) ? JoinPinComponent(PinNamePrefix, pinName, "Vih", suffixText) : JoinPinComponent(PinNamePrefix, pinName, "Vih", concurrentSuffix, suffixText), GetSelectorList());
            var specVil = new DcSpec(string.IsNullOrEmpty(concurrentSuffix) ? JoinPinComponent(PinNamePrefix, pinName, "Vil", suffixText) : JoinPinComponent(PinNamePrefix, pinName, "Vil", concurrentSuffix, suffixText), GetSelectorList());
            var specVoh = new DcSpec(string.IsNullOrEmpty(concurrentSuffix) ? JoinPinComponent(PinNamePrefix, pinName, "Voh", suffixText) : JoinPinComponent(PinNamePrefix, pinName, "Voh", concurrentSuffix, suffixText), GetSelectorList());
            var specVol = new DcSpec(string.IsNullOrEmpty(concurrentSuffix) ? JoinPinComponent(PinNamePrefix, pinName, "Vol", suffixText) : JoinPinComponent(PinNamePrefix, pinName, "Vol", concurrentSuffix, suffixText), GetSelectorList());
            var specVt = new DcSpec(string.IsNullOrEmpty(concurrentSuffix) ? JoinPinComponent(PinNamePrefix, pinName, "Vt", suffixText) : JoinPinComponent(PinNamePrefix, pinName, "Vt", concurrentSuffix, suffixText), GetSelectorList());
            if (ioInfoRow.IsUsePowerPinValue && !_testSettings.PowerPinList.Exists(s => s.Equals(ioInfoRow.Nv, StringComparison.OrdinalIgnoreCase)))
            {
                Response.Report($"'{ioInfoRow.Nv}' used in IoInfo sheet, but missing in TestSetting sheet!", EnumMessageLevel.Error, 45);
            }
            foreach (string categoryName in list)
            {
                specVih.AddCategory(CreateClassifiedCategoryItem("Vih", ioInfoRow, categoryName, suffix));
                specVil.AddCategory(CreateClassifiedCategoryItem("Vil", ioInfoRow, categoryName, suffix));
                specVoh.AddCategory(CreateClassifiedCategoryItem("Voh", ioInfoRow, categoryName, suffix));
                specVol.AddCategory(CreateClassifiedCategoryItem("Vol", ioInfoRow, categoryName, suffix));
                specVt.AddCategory(CreateClassifiedCategoryItem("Vt", ioInfoRow, categoryName, suffix));
            }
            result.Add(specVih);
            result.Add(specVil);
            result.Add(specVoh);
            result.Add(specVol);
            result.Add(specVt);

            return result;
        }

        internal List<Selector> GetSelectorList()
        {
            var result = new List<Selector>
            {
                new Selector("Min", "Min"), new Selector("Typ", "Typ"), new Selector("Max", "Max")
            };
            return result;
        }

        private CategoryInSpec CreateClassifiedCategoryItem(string suffix, IoInfoRow ioInfoRow, string categoryName, string blockSuffix = "")
        {
            string typ, min, max;
            string pinName = ioInfoRow.PinGrpName;
            var categoryInfo = new DcCategoryInfo(categoryName);

            if (categoryInfo.Test.Equals("Conti", StringComparison.OrdinalIgnoreCase) && (_ioInfo.GetBlockPins("DCTEST_Continuity").Contains(pinName) || (_ioInfoConcurrent != null && _ioInfoConcurrent.GetBlockPins("DCTEST_Continuity").Contains(pinName))))
            {
                if (ioInfoRow.IsUsePowerPinValue)
                {
                    if (string.IsNullOrEmpty(blockSuffix))
                    {
                        typ = "_" + JoinPinComponent(PinNamePrefix, ioInfoRow.Nv, "VAR") + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Ratio", suffix, "Conti", GlobalSpecSuffix);
                    }
                    else
                    {
                        typ = "_" + JoinPinComponent(PinNamePrefix, ioInfoRow.Nv, "VAR_" + blockSuffix) + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Ratio", suffix, "Conti", GlobalSpecSuffix);
                    }

                    min = typ;
                    max = typ;
                }
                else
                {
                    typ = "_" + JoinPinComponent(PinNamePrefix, pinName, suffix, "Conti", GlobalSpecSuffix);
                    min = typ + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Conti", GlobalSpecSuffix, "Minus");
                    max = typ + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Conti", GlobalSpecSuffix, "Plus");
                }
            }
            else if (categoryInfo.Test.Equals("HardIP", StringComparison.OrdinalIgnoreCase)
                && (_ioInfo.GetBlockPins("HardIP").Contains(pinName)
                    || (_ioInfoConcurrent != null && _ioInfoConcurrent.GetBlockPins("HardIP").Contains(pinName))))
            {
                if (ioInfoRow.IsUsePowerPinValue)
                {
                    if (string.IsNullOrEmpty(blockSuffix))
                    {
                        typ = "_" + JoinPinComponent(PinNamePrefix, ioInfoRow.Nv, "VAR") + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Ratio", suffix, "H", GlobalSpecSuffix);
                    }
                    else
                    {
                        typ = "_" + JoinPinComponent(PinNamePrefix, ioInfoRow.Nv, "VAR_" + blockSuffix) + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Ratio", suffix, "H", GlobalSpecSuffix);
                    }

                    min = typ;
                    max = typ;
                }
                else
                {
                    typ = "_" + JoinPinComponent(PinNamePrefix, pinName, suffix, "H", GlobalSpecSuffix);
                    min = typ + "*_" + JoinPinComponent(PinNamePrefix, pinName, "H", GlobalSpecSuffix, "Minus");
                    max = typ + "*_" + JoinPinComponent(PinNamePrefix, pinName, "H", GlobalSpecSuffix, "Plus");
                }
            }
            else if (Regex.IsMatch(categoryInfo.Test, "(Sa|Td|SAChain|PDF)", RegexOptions.IgnoreCase)
                && (_ioInfo.GetBlockPins("Scan").Contains(pinName)
                    || (_ioInfoConcurrent != null && _ioInfoConcurrent.GetBlockPins("Scan").Contains(pinName))))
            {
                if (ioInfoRow.IsUsePowerPinValue)
                {
                    if (string.IsNullOrEmpty(blockSuffix))
                    {
                        typ = "_" + JoinPinComponent(PinNamePrefix, ioInfoRow.Nv, "VAR") + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Ratio", suffix, "Scan", GlobalSpecSuffix);
                    }
                    else
                    {
                        typ = "_" + JoinPinComponent(PinNamePrefix, ioInfoRow.Nv, "VAR_" + blockSuffix) + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Ratio", suffix, "Scan", GlobalSpecSuffix);
                    }

                    min = typ;
                    max = typ;
                }
                else
                {
                    typ = "_" + JoinPinComponent(PinNamePrefix, pinName, suffix, "Scan", GlobalSpecSuffix);
                    min = typ + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Scan", GlobalSpecSuffix, "Minus");
                    max = typ + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Scan", GlobalSpecSuffix, "Plus");
                }
            }
            else
            {
                if (ioInfoRow.IsUsePowerPinValue)
                {
                    if (string.IsNullOrEmpty(blockSuffix))
                    {
                        typ = "_" + JoinPinComponent(PinNamePrefix, ioInfoRow.Nv, "VAR") + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Ratio", suffix, GlobalSpecSuffix);
                    }
                    else
                    {
                        typ = "_" + JoinPinComponent(PinNamePrefix, ioInfoRow.Nv, "VAR_" + blockSuffix) + "*_" + JoinPinComponent(PinNamePrefix, pinName, "Ratio", suffix, GlobalSpecSuffix);
                    }

                    min = typ;
                    max = typ;
                }
                else
                {
                    typ = "_" + JoinPinComponent(PinNamePrefix, pinName, suffix, GlobalSpecSuffix);
                    min = typ + "*_" + JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix, "Minus");
                    max = typ + "*_" + JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix, "Plus");
                }
                if (!ioInfoRow.Type.Equals("Common", StringComparison.OrdinalIgnoreCase))
                {
                    typ = "0";
                    min = "0";
                    max = "0";
                }
            }

            if (categoryName.Equals(_testSettings.DcCategoryInfos.FindContiCategory(out EnumMessageLevel _, out string _), StringComparison.OrdinalIgnoreCase))
            {
                typ = "0";
                min = "0";
                max = "0";
            }
            return new CategoryInSpec(categoryName, FormulaPrefix + typ, FormulaPrefix + min, FormulaPrefix + max);
        }

        internal string MakeCeilingFormula(string cell)
        {
            string result;
            double.TryParse(LocalSpecs.PwrSupplyRes, out double value);
            if (value >= 0.001)
            {
                result = "CEILING(" + cell + "," + value.ToString("G15", CultureInfo.InvariantCulture) + ")";
            }
            else
            {
                result = cell;
            }

            return result;
        }

        internal string MakeFloorFormula(string cell)
        {
            string result;
            double.TryParse(LocalSpecs.PwrSupplyRes, out double value);
            if (value >= 0.001)
            {
                result = "FLOOR(" + cell + "," + value.ToString("G15", CultureInfo.InvariantCulture) + ")";
            }
            else
            {
                result = cell;
            }

            return result;
        }
    }
}
