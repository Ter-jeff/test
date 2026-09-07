using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class AcCategoryGenerator
    {
        private AcSpecSheet _acSpecSheet;
        private List<AcSpec> _varList = new List<AcSpec>();
        public void GenAcCategory(Dictionary<string, HardIpSheet> planDic, List<InstanceSheet> instSheets)
        {
            if (TestProgram.IgxlWorkBk.AcSpecSheets.Count == 0)
            {
                return;
            }

            _acSpecSheet = TestProgram.IgxlWorkBk.AcSpecSheets.Values.ToArray()[0];
            _varList = _acSpecSheet.Rows.ToList();
            AddDefaultSpecs();

            var instsheetAll = new List<InstanceRow>();
            foreach (InstanceSheet instsheet in instSheets)
            {
                instsheetAll.AddRange(instsheet.Rows);
            }

            foreach (string sheet in planDic.Keys)
            {
                foreach (HardIpPattern pattern in planDic[sheet].Rows)
                {
                    string timeSets = string.Empty;
                    string timeSet2Cat = string.Empty;

                    if (!string.IsNullOrEmpty(pattern.AcUsed))
                    {
                        InstanceRow inst = instsheetAll.FirstOrDefault(y => y.Args[0].Equals(pattern.Pattern.GetLastPayload(), StringComparison.OrdinalIgnoreCase));
                        if (inst == null)
                        {
                            continue;
                        }

                        timeSets = inst.TimeSets;
                        timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets, BlockType.HardIp);
                        if (timeSet2Cat == "TBD")
                        {
                            timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets);
                        }

                        List<Timing> timings = pattern.GetTimingsByAc();
                        string blockName = pattern.SheetName.ToUpper().Replace("HARDIP_", "").Replace(" ", "").Replace("_", "");
                        string categoryName = "";
                        if (timeSet2Cat == "")
                        {
                            categoryName = blockName + "_" + timeSet2Cat + timings.Aggregate("", (current, timing) =>
                                current + timing.Name + "_" + timing.SuffixAcSpecName + "_").Trim('_');
                        }
                        else
                        {
                            categoryName = blockName + "_" + timeSet2Cat + "_" + timings.Aggregate("", (current, timing) =>
                                current + timing.Name + "_" + timing.SuffixAcSpecName + "_").Trim('_');
                        }

                        AddAcCategory(timings, categoryName, timeSet2Cat);
                    }
                }
            }
        }

        private void AddAcCategory(List<Timing> timings, string categoryName, string acCategory)
        {
            if (_acSpecSheet.CategoryList.Contains(categoryName))
            {
                return;
            }

            _acSpecSheet.CategoryList.Add(categoryName);
            foreach (AcSpec specs in _varList)
            {
                string pinName = specs.Symbol.Replace("_Freq_VAR", "");
                if (specs.Symbol != "")
                {
                    string type;
                    string min;
                    string max;

                    #region get default
                    if (specs.ContainsCategory(acCategory))
                    {
                        int index = 0;
                        for (int cnt = 0; cnt < specs.CategoryList.Count; cnt++)
                        {
                            if (specs.CategoryList[cnt].Name.Equals(acCategory))
                            {
                                index = cnt;
                                break;
                            }
                        }
                        type = specs.CategoryList[index].Typ;
                        min = specs.CategoryList[index].Min;
                        max = specs.CategoryList[index].Max;
                    }
                    else
                    {
                        type = specs.CategoryList[0].Typ;
                        min = specs.CategoryList[0].Min;
                        max = specs.CategoryList[0].Max;
                    }
                    #endregion

                    foreach (Timing timing in timings)
                    {
                        if (timing.Name.Equals(pinName, StringComparison.OrdinalIgnoreCase))
                        {
                            type = timing.Type != "" ? timing.Type : type;
                            min = timing.Min != "" ? timing.Min : min;
                            max = timing.Max != "" ? timing.Max : max;
                        }
                    }

                    var newCategory = new CategoryInSpec(categoryName, type, min, max);
                    specs.AddCategory(newCategory);
                }
            }
        }

        private void AddDefaultSpecs()
        {
            List<ProtocolAwarePin> defaultTimings = NwireSingleton.Instance().SettingInfo.NwirePins;
            var lVarList = new List<string>();
            foreach (AcSpec data in _varList)
            {
                lVarList.Add(data.Symbol);
            }
            foreach (ProtocolAwarePin timing in defaultTimings)
            {
                if (lVarList.Contains(timing.CreatePinNameWithDiff() + "_Freq_VAR"))
                {
                    continue;
                }
                var selectorList = new List<Selector>
                {
                    new Selector("Typ", "Typ"), new Selector("Min", "Min"), new Selector("Max", "Max")
                };
                var acSpecs = new AcSpec(timing.CreatePinNameWithDiff() + "_Freq_VAR", selectorList);

                List<string> acFullCategoryList = _acSpecSheet.CategoryList;
                // Write Category
                foreach (string categroyName in acFullCategoryList)
                {
                    string value = "= _" + timing.CreatePinNameWithDiff() + "_Freq_GLB";
                    var categroy = new CategoryInSpec(categroyName, value, value, value);
                    acSpecs.AddCategory(categroy);
                }

                _acSpecSheet.AddRow(acSpecs);
            }
        }
    }
}
