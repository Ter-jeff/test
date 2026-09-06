using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using LogLib.Static;

using ScghLib.Reader;

using TestPlanLib.Basic;

using static Automation.GenerateIgxl.BistBira.NonLogicData.BistNonLogicalLib;

namespace Automation.GenerateIgxl.BinCut.Business
{
    internal class BinCutAcSpecsWriter
    {
        private readonly Dictionary<string, PatternData> _patternDatas = AcTSetCategoryMapSingleton.Instance().PatternList;
        internal TimeSettingSheet _timeSettingSheet = TestPlanStatic.TimeSettingSheet;
        internal AcSpecSheet _acSpecSheet;
        internal readonly HashSet<string> _alreadySetupAc = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public void GenAcSpecs(List<BinCutFinalInstanceRow> binCutFinalInstanceRows, AcSpecSheet acSpecSheet, bool isEvs = false)
        {
            _acSpecSheet = acSpecSheet;

            foreach (BinCutFinalInstanceRow binCutFinalInstanceRow in binCutFinalInstanceRows)
            {
                string shiftSpeed = binCutFinalInstanceRow.BinCutInstanceRow.ShiftSpeed;
                string shiftSymbol = "";
                string newTimeSet = binCutFinalInstanceRow.BinCutInstanceRow.TimeSet;
                string userDefineAc = "";
                if (binCutFinalInstanceRow.BinCutInstanceRow.TimeSet.Contains(":"))
                {
                    newTimeSet = binCutFinalInstanceRow.BinCutInstanceRow.TimeSet.Split(':').First();
                    userDefineAc = binCutFinalInstanceRow.BinCutInstanceRow.TimeSet.Split(':').Last();
                }
                if (string.IsNullOrEmpty(shiftSpeed) && string.IsNullOrEmpty(newTimeSet))
                {
                    continue;
                }

                string timeSetVersion = binCutFinalInstanceRow.GetTimeSetVersion(binCutFinalInstanceRow.PatternList);

                string acSpecName = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSetVersion);
                if (string.IsNullOrEmpty(acSpecName) || acSpecName.ToUpper().Equals("TBD"))
                {
                    Response.Report($"Can not find AC spec for {timeSetVersion} ...", EnumMessageLevel.Error);
                    continue;
                    //acSpecName = GetCategoryFromTimeSet(timeSets);
                }

                if (!acSpecName.ToUpper().StartsWith("TBD"))
                {
                    if (_patternDatas != null && _patternDatas.Any())
                    {
                        shiftSymbol = UpdateSymbolByTimeSet(_patternDatas, timeSetVersion, binCutFinalInstanceRow.PayloadList);
                    }

                    if (isEvs)
                    {
                        shiftSymbol = "ShiftIn_Freq_VAR";
                    }

                    if (!string.IsNullOrEmpty(shiftSpeed))
                    {
                        if (string.IsNullOrEmpty(shiftSymbol))
                        {
                            Response.Report($"Can not get any variables name in the category {userDefineAc}...", EnumMessageLevel.Error);
                        }
                        else if (!acSpecSheet.Rows.Any(x => x.Symbol.Equals(shiftSymbol, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            Response.Report($"Can not find variables {shiftSymbol} in AC spec ...", EnumMessageLevel.Error);
                        }
                    }

                }
                binCutFinalInstanceRow.BinCutInstanceRow.TimeSet = newTimeSet;
                binCutFinalInstanceRow.BinCutInstanceRow.AcSpec = UpdateAcSpecByShift(acSpecName, timeSetVersion, shiftSpeed, shiftSymbol, userDefineAc);
            }
        }

        public void GenAcSpecs(List<BistProdFlowRow> bistProdFlowRows, AcSpecSheet acSpecSheet)
        {
            _acSpecSheet = acSpecSheet;

            var genInstanceRow = new GenInstanceRow();
            foreach (BistProdFlowRow bistProdFlowRow in bistProdFlowRows)
            {
                if (bistProdFlowRow.TimeSet.Contains(":"))
                {
                    string[] parts = bistProdFlowRow.TimeSet.Split(':');
                    if (parts.Length < 2)
                    {
                        return;
                    }

                    string newTimeSet = parts[0];
                    string userDefineAc = parts[1];

                    string timeSet = genInstanceRow.GetTimeSetVersion(newTimeSet);
                    string acSpecName = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSet);

                    UpdateAcSpecByShift(acSpecName, timeSet, "", "", userDefineAc);
                }
            }
        }

        internal string UpdateAcSpecByShift(string originalAc, string timeSetVersion, string shiftFreq, string shiftSymbol, string userDefinedAc)
        {
            string newAcName = originalAc;

            if (!string.IsNullOrEmpty(userDefinedAc))
            {
                newAcName = userDefinedAc;
                if (!_alreadySetupAc.Contains(newAcName))
                {
                    if (_timeSettingSheet == null)
                    {
                        Response.Report("Can not find TimeSetting sheet in TestPlan ...", EnumMessageLevel.Error);
                        return "";
                    }

                    TimeSettingRow timeSettings = _timeSettingSheet.Rows.Find(x => x.SheetName.Equals(userDefinedAc));

                    if (timeSettings == null)
                    {
                        Response.Report($"Can not find AC spec {userDefinedAc} in TimeSetting sheet ...", EnumMessageLevel.Error);
                        return "Missing user define AC";
                    }
                    bool category = _acSpecSheet.CategoryList.Exists(x => x.Equals(newAcName, StringComparison.OrdinalIgnoreCase));

                    if (!category)
                    {
                        _acSpecSheet.CategoryList.Add(newAcName);
                    }

                    foreach (AcSpec acSpec in _acSpecSheet.Rows)
                    {
                        if (string.IsNullOrEmpty(acSpec.Symbol))
                        {
                            continue;
                        }

                        string foundSymbol = timeSettings.SymbolValues.Keys.FirstOrDefault(x => x.Equals(acSpec.Symbol, StringComparison.OrdinalIgnoreCase));

                        if (!string.IsNullOrEmpty(foundSymbol))
                        {
                            string value = timeSettings.SymbolValues[foundSymbol].ConvertNumber();
                            if (category)
                            {
                                CategoryInSpec target = acSpec.CategoryList.Find(x => x.Name.Equals(originalAc, StringComparison.OrdinalIgnoreCase));
                                target.Max = value;
                                target.Min = value;
                                target.Typ = value;
                            }
                            else
                            {
                                acSpec.AddCategory(ModifyValueFromTimeSettings(originalAc, newAcName, value, acSpec));
                            }
                        }
                        else if (acSpec.Symbol.Equals("Using TSet", StringComparison.CurrentCultureIgnoreCase))
                        {
                            if (!category)
                            {
                                acSpec.AddCategory(ModifyValueFromTimeSettings(originalAc, newAcName, timeSetVersion, acSpec));
                            }
                        }
                        else
                        {
                            if (!category)
                            {
                                acSpec.AddCategory(ModifyValueFromTimeSettings(originalAc, newAcName, "", acSpec));
                            }
                        }
                    }
                    _alreadySetupAc.Add(newAcName);
                }
            }
            if (!string.IsNullOrEmpty(shiftFreq))
            {
                string shiftAcName = $"{newAcName}_{shiftFreq}";
                if (!_alreadySetupAc.Contains(shiftAcName) && !string.IsNullOrEmpty(shiftSymbol))
                {
                    bool category = _acSpecSheet.CategoryList.Exists(x => x.Equals(shiftAcName, StringComparison.OrdinalIgnoreCase));

                    if (!category)
                    {
                        _acSpecSheet.CategoryList.Add(shiftAcName);
                    }

                    foreach (AcSpec acSpec in _acSpecSheet.Rows)
                    {
                        if (string.IsNullOrEmpty(acSpec.Symbol))
                        {
                            continue;
                        }

                        if (acSpec.Symbol.Equals(shiftSymbol, StringComparison.OrdinalIgnoreCase))
                        {
                            string value = shiftFreq.ConvertNumber();
                            if (category)
                            {
                                CategoryInSpec target = acSpec.CategoryList.Find(x => x.Name.Equals(shiftAcName, StringComparison.OrdinalIgnoreCase));
                                target.Max = value;
                                target.Min = value;
                                target.Typ = value;
                            }
                            else
                            {
                                acSpec.AddCategory(ModifyValueFromTimeSettings(newAcName, shiftAcName, value, acSpec));
                            }
                        }
                        else if (acSpec.Symbol.Equals("Using TSet", StringComparison.CurrentCultureIgnoreCase))
                        {
                            if (!category)
                            {
                                acSpec.AddCategory(ModifyValueFromTimeSettings(newAcName, shiftAcName, timeSetVersion, acSpec));
                            }
                        }
                        else
                        {
                            if (!category)
                            {
                                acSpec.AddCategory(ModifyValueFromTimeSettings(newAcName, shiftAcName, "", acSpec));
                            }
                        }
                    }
                    _alreadySetupAc.Add(shiftAcName);
                }
                newAcName = shiftAcName;
            }
            return newAcName;
        }

        internal CategoryInSpec ModifyValueFromTimeSettings(string originalAcSpec, string acSpecName, string value, AcSpec acSpec)
        {
            if (!string.IsNullOrEmpty(value))
            {
                CategoryInSpec item = acSpec.Symbol.Equals("Using TSet", StringComparison.CurrentCultureIgnoreCase) ? new CategoryInSpec(acSpecName, value, "", "") : new CategoryInSpec(acSpecName, value, value, value);
                return item;
            }

            if (acSpec.CategoryList.Exists(x => x.Name.Equals(originalAcSpec, StringComparison.CurrentCultureIgnoreCase)))
            {
                CategoryInSpec row = acSpec.CategoryList.Find(x => x.Name.Equals(originalAcSpec, StringComparison.CurrentCultureIgnoreCase));
                var item = new CategoryInSpec(acSpecName, row.Typ, row.Min, row.Max);
                return item;
            }
            else
            {
                CategoryInSpec row = acSpec.CategoryList.First();
                var item = new CategoryInSpec(acSpecName, row.Typ, row.Min, row.Max);
                return item;
            }
        }

        internal string UpdateSymbolByTimeSet(Dictionary<string, PatternData> patternDatas, string timeSetVersion, List<string> payloadList)
        {
            List<ComTimeSetBasicSheet> timeSetSheets = AcTSetCategoryMapSingleton.Instance().MultiTimeSetSheets;
            var timeSetInfo = timeSetSheets.Where(x => x.Name.Equals(timeSetVersion)).ToList();
            if (!timeSetInfo.Any())
            {
                return "";
            }

            foreach (string pattern in payloadList)
            {
                if (!patternDatas.ContainsKey(pattern.ToLower()))
                {
                    continue;
                }

                string scanTSet = patternDatas[pattern.ToLower()].ScanTset;
                if (!string.IsNullOrEmpty(scanTSet))
                {
                    IEnumerable<TSet> periodInfo = timeSetInfo.First().Rows.Where(x => x.Name.Equals(scanTSet, StringComparison.CurrentCultureIgnoreCase)).ToList();
                    if (periodInfo.Any())
                    {
                        return periodInfo.First().CyclePeriod.Split(new[] { '/', ' ', '(', ')' }, StringSplitOptions.RemoveEmptyEntries).Last().TrimStart('_');
                    }
                }
            }
            return "";
        }
    }
}
