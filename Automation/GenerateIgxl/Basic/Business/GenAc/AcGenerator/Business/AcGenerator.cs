using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenAc.AcInput.BassData;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.Reader;
using Automation.Reader.ConfigFile.TimingFileCategoryMapping;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.PatternListManager;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;
using IgxlLib.Utility;

using TestPlanLib.Basic;
using TestPlanLib.Const;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.Basic.Business.GenAc.AcGenerator.Business
{
    public class AcGenerator
    {
        #region Field
        public const string Scan = "Scan";
        public const string Mbist = "Mbist";
        public string Rtos = ModuleSingleton.Instance().ModuleRtos;
        public const string SpiRom = "SPI_ROM";
        public const string Common = "Common";
        public const string HardIp = "HardIP";
        public const string Jtag = "JTAG";

        private const string SpiRomCat = "SPI_ROM";
        private const string TckFreqVar = "TCK_Freq_VAR";

        private readonly TimingFileCategoryMappingConfig _timingFIleCategoryConfig;
        private AcSpecSheet _acSpecSheet;
        protected AcInputSheet AcInputSheet;
        internal List<string> _acFullCategoryList;

        #endregion

        #region Constructor
        public AcGenerator(AcInputSheet acInputSheet, AcSpecSheet acSpecSheet = null)
        {
            AcInputSheet = acInputSheet;
            _acSpecSheet = acSpecSheet;
            if (SettingStatic.BasicConfigWorkbook != null)
            {
                _timingFIleCategoryConfig = TimingFileCategoryMappingConfig.LoasConfig(SettingStatic.BasicConfigWorkbook.Worksheets[CommonConst.TimingFileCategoryMappingSheetName]);
            }
        }
        #endregion

        #region Function
        public AcSpecSheet GenerateFlow(TimeSetSheets timeSetSheets, List<PatternData> patlists)
        {
            _acFullCategoryList = InitialAcCatList();

            var selectorNameList = new List<string> { "Typ", "Min", "Max" };
            _acSpecSheet = new AcSpecSheet("AC_Specs", _acFullCategoryList, selectorNameList);

            WriteAcData();

            if (patlists != null)
            {
                AcSpecSheetUpdateEquation(timeSetSheets, patlists);
            }

            AddComment();

            UpdateAcSpecByTimesettings();
            return _acSpecSheet;
        }


        private void UpdateAcSpecByTimesettings()
        {
            if (TestPlanStatic.TimeSettingSheet == null)
            {
                return;
            }

            TimeSettingSheet timesettingsheet = TestPlanStatic.TimeSettingSheet;
            foreach (TimeSettingRow timeSettings in timesettingsheet.Rows)
            {
                string timeSetVersion = AcTSetCategoryMapSingleton.Instance().GetTimeSetVersion(timeSettings.SheetName);
                if (string.IsNullOrEmpty(timeSetVersion))
                {
                    continue;
                }

                string acSpecName = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSetVersion);
                if (string.IsNullOrEmpty(acSpecName))
                {
                    continue;
                }

                if (_acSpecSheet.CategoryList.Exists(x => x.Equals(acSpecName)))
                {
                    List<string> allExistCateogry = _acSpecSheet.CategoryList.FindAll(x => x.Equals(acSpecName));
                    foreach (string category in allExistCateogry)
                    {
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
                                CategoryInSpec target = acSpec.CategoryList.Find(x => x.Name.Equals(category, StringComparison.OrdinalIgnoreCase));
                                target.Max = value;
                                target.Min = value;
                                target.Typ = value;
                            }
                            else if (acSpec.Symbol.Equals("Using TSet", StringComparison.CurrentCultureIgnoreCase))
                            {
                                CategoryInSpec target = acSpec.CategoryList.Find(x => x.Name.Equals(category, StringComparison.OrdinalIgnoreCase));
                                target.Min = timeSetVersion;
                            }
                        }
                    }
                }
            }
        }

        public void AcSpecSheetUpdateEquation(TimeSetSheets timeSetSheets, List<PatternData> patlists)
        {
            IOrderedEnumerable<ComTimeSetBasicSheet> timeSetSheetBySort = timeSetSheets.OrderBy(x => x.Name);
            foreach (ComTimeSetBasicSheet timeSetBasicSheet in timeSetSheetBySort)
            {
                var lVarValueDict = new Dictionary<string, double>();
                // get TimeSetSheet using TSet VarValue Map Dict
                foreach (ComTimeSetBasicSheet.TSetEqnVarMap tSetEquObject in timeSetBasicSheet.AllTSetEqnVariable)
                {
                    foreach (KeyValuePair<string, double> variable in tSetEquObject.DictVariable)
                    {
                        if (!lVarValueDict.ContainsKey(variable.Key.ToUpper()))
                        {
                            lVarValueDict.Add(variable.Key.ToUpper(), variable.Value);
                        }
                    }
                }

                double lVarValueDictTck = lVarValueDict.Where(entry => entry.Key == "TCK_FREQ_VAR").Select(entry => entry.Value).FirstOrDefault();
                string acCategoryName = GetTimesetcategory(timeSetBasicSheet.Name, lVarValueDictTck);     //TimeSet Sheet default block name, ex: SocMbist/GfxScan
                BlockType blockTypeName = GetTimesetBlockType(timeSetBasicSheet.Name);     //only could be one of BlockType.Mbist/BlockType.Scan/Block.HardIp
                if (lVarValueDict.Count > 0)
                {

                    UpdateSymbolSpecs(lVarValueDict);
                    acCategoryName = TSetMapCategory(lVarValueDict, acCategoryName);
                }
                else
                {
                    if (blockTypeName == BlockType.None)
                    {
                        acCategoryName = "Common";
                        blockTypeName = BlockType.Common;
                    }
                }

                if (!AcTSetCategoryMapSingleton.Instance().Contains(timeSetBasicSheet.Name))
                {
                    AcTSetCategoryMapSingleton.Instance().SetRow(timeSetBasicSheet.Name, blockTypeName, acCategoryName);
                }
            }

            GenEfuseACspec(patlists);
        }

        private void AddComment()
        {
            var categoryTimingSheet = new List<Tuple<string, string>>();
            foreach (string category in _acSpecSheet.CategoryList)
            {
                string timingsheetname = AcTSetCategoryMapSingleton.Instance().GetCategoryUsageTSetSheetName(category);
                categoryTimingSheet.Add(new Tuple<string, string>(category, timingsheetname));
            }
            var selectorList = new List<Selector> { new Selector("", "") };

            _acSpecSheet.AddRow(new AcSpec("", selectorList));
            _acSpecSheet.AddRow(new AcSpec("", selectorList));
            var commentRow = new AcSpec("Using TSet", selectorList);
            foreach (Tuple<string, string> timingsheetname in categoryTimingSheet)
            {
                var item = new CategoryInSpec(timingsheetname.Item1, timingsheetname.Item2, "", "");
                commentRow.AddCategory(item);
            }
            // timing mapping line
            _acSpecSheet.AddRow(commentRow);
        }

        #region Virtual Methods
        protected virtual void GenEfuseACspec(List<PatternData> patlists)
        {
            int efuseIndex = 0;
            var lHardIp2EfuseDict = new Dictionary<string, string>();
            //Collect all timeSetSheets that are used by efuse pattern
            IOrderedEnumerable<string> lEfuseTSetSheets = PatternListEfuseTimeSet.GetEfusePatternTimeSet(patlists).OrderBy(x => x);
            foreach (string efuseTSet in lEfuseTSetSheets) //all timeSetSheets that are used by efuse pattern
            {
                string efuse = "Efuse";
                if (!AcTSetCategoryMapSingleton.Instance().Contains(efuseTSet))
                {
                    continue; // error 
                }
                string hardipCate = AcTSetCategoryMapSingleton.Instance().GetCategory(efuseTSet);
                if (!lHardIp2EfuseDict.TryGetValue(hardipCate, out string value))
                {
                    efuse = efuseIndex == 0 ? efuse : efuse + "_" + efuseIndex;
                    CopyAcCategory(hardipCate, efuse);
                    lHardIp2EfuseDict.Add(hardipCate, efuse);
                    AcTSetCategoryMapSingleton.Instance().SetRow(efuseTSet, BlockType.Efuse, efuse); //!! add efuse timeset to category!!
                    efuseIndex++;
                }
                else // already create an efuse category
                {
                    AcTSetCategoryMapSingleton.Instance().SetRow(efuseTSet, BlockType.Efuse, value); //!! use the exsis efuse cat
                }
            }
        }

        protected virtual List<string> InitialAcCatList()
        {
            var categoryList = new List<string>
            {
                //Common
                Common, 
                //CpuScan and CpuMbist
                ModuleSingleton.Instance().ModuleCpu + Scan, ModuleSingleton.Instance().ModuleCpu + Mbist,
                //GfxScan and GfxMbist
                ModuleSingleton.Instance().ModuleGfx + Scan,
                ModuleSingleton.Instance().ModuleGfx + Mbist,
                //SocScan and SocMbist
                ModuleSingleton.Instance().ModuleSoc + Scan,
                ModuleSingleton.Instance().ModuleSoc + Mbist,
                //SPI_ROM
                SpiRom,
                //HardIP
                HardIp
            };

            DataTable table = NwireSingleton.Instance().SettingInfo.SettingTable;
            for (int j = 0; j < table.Rows.Count; j++)
            {
                string value = table.Rows[j][0].ToString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    categoryList.Add($"{Common}_{value}");
                }
            }

            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                List<string> names = NwireSingleton.Instance().ReferenceFlow();
                foreach (string name in names)
                {
                    if (!categoryList.Exists(p => p.Equals(name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        categoryList.Add(name);
                    }
                }
            }

            return categoryList.Distinct().ToList();
        }

        protected virtual string GetTimesetcategory(string sheetName, double lVarValueDictTck)
        {
            long lVarValueDictTckLon = (long)(lVarValueDictTck / Math.Pow(10, 6));
            // Config assign ac category 
            string timingCategory = "";
            if (_timingFIleCategoryConfig != null)
            {
                timingCategory = _timingFIleCategoryConfig.GetCategoryMapping(sheetName);
            }

            if (!string.IsNullOrEmpty(timingCategory))
            {
                return timingCategory;
            }

            //ex: TIMESET_OSPA0_A_AN_SI_1   
            string[] toks = sheetName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            string moduleBlock = HardIpConstData.HardIp; //Soc/Gfx/Cpu/HardIp/Other...... default set to HardIp
            if (toks.Length >= 5)
            {
                //Main block part
                string moduleCpu = ModuleSingleton.Instance().ModuleCpu;
                string moduleGfx = ModuleSingleton.Instance().ModuleGfx;
                string moduleSoc = ModuleSingleton.Instance().ModuleSoc;
                switch (toks[2].ToUpper()[0])
                {
                    case 'S':
                    case 'H':
                        moduleBlock = moduleSoc;
                        break;
                    case 'C':
                        moduleBlock = moduleCpu;
                        break;
                    case 'L':
                        moduleBlock = moduleGfx;
                        break;
                }

                BlockType blockType = GetBlockType(toks);  //Scan/Mbist/Other
                if (blockType == BlockType.None)
                {
                    return "AC_" + sheetName;
                }

                if (blockType == BlockType.HardIp)
                {
                    return HardIpConstData.HardIp + $"_{lVarValueDictTckLon}MHz";
                }

                if (blockType == BlockType.SPI_ROM)
                {
                    return BlockType.SPI_ROM.ToString();
                }


                if (moduleBlock == HardIpConstData.HardIp || blockType.ToString() == HardIpConstData.HardIp)
                {
                    return HardIpConstData.HardIp;
                }



                return moduleBlock + blockType;
            }
            return HardIpConstData.HardIp;
        }


        protected internal virtual BlockType GetTimesetBlockType(string sheetName)
        {
            BlockType blockType = BlockType.HardIp;
            string[] toks = sheetName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (toks.Length >= 5)
            {
                blockType = GetBlockType(toks);
            }

            return blockType;
        }

        public static BlockType GetBlockType(string[] toks)
        {
            if (toks.ToList().Exists(x => x.Equals("Scan", StringComparison.CurrentCultureIgnoreCase) ||
                x.Equals("Saa", StringComparison.CurrentCultureIgnoreCase)))
            {
                return BlockType.Scan;
            }
            if (toks.ToList().Exists(x => x.Equals("Rtos", StringComparison.CurrentCultureIgnoreCase)))
            {
                return BlockType.SPI_ROM;
            }
            string block = toks[3];
            if (block.Equals("AN", StringComparison.CurrentCultureIgnoreCase))
            {
                return BlockType.HardIp;
            }

            if (block.Equals("JT", StringComparison.CurrentCultureIgnoreCase))
            {
                return BlockType.HardIp;
            }

            if (block.Equals("SC", StringComparison.CurrentCultureIgnoreCase))
            {
                return BlockType.Scan;
            }

            if (block.Equals("BI", StringComparison.CurrentCultureIgnoreCase))
            {
                return BlockType.Mbist;
            }

            if (toks.ToList().Exists(x => x.Equals("bsr", StringComparison.CurrentCultureIgnoreCase) ||
                                          x.Equals("mbist", StringComparison.CurrentCultureIgnoreCase)))
            {
                return BlockType.Mbist;
            }

            return BlockType.None;
        }

        #endregion

        private void WriteAcData()
        {
            if (AcInputSheet != null)
            {
                foreach (AcInputRow acInputRow in AcInputSheet.AcInputData)
                {
                    WriteOneRecordToAcSpecSheet(acInputRow.Symbol, acInputRow.Value, acInputRow.Typ, acInputRow.Min, acInputRow.Max);
                }
                WriteJitterSymbolToAcSpecSheet();
            }
        }

        private void WriteJitterSymbolToAcSpecSheet()
        {
            AcSpec acSpecs = WriteOneRecordToAcSpec("J_TargetResolution", "1e-12", "1e-12", "1e-12", "1e-12");
            _acSpecSheet.AddRow(acSpecs);

            acSpecs = WriteOneRecordToAcSpec("J_RepeatTimes", "500", "500", "500", "500");
            _acSpecSheet.AddRow(acSpecs);
        }

        internal void WriteOneRecordToAcSpecSheet(string pStrSymbol, string pStrValue, string pStrTyp, string pStrMin, string pStrMax)
        {
            List<ProtocolAwarePin> nWirePin = NwireSingleton.Instance().SettingInfo.NwirePins;
            int index = nWirePin.FindIndex(s =>
                    !string.IsNullOrEmpty(s.OutClk) &&
                    !string.IsNullOrEmpty(pStrSymbol) &&
                    pStrSymbol.IndexOf(s.OutClk, StringComparison.OrdinalIgnoreCase) >= 0
                );

            string acSpecSymbol = GetAcSpecSymbol(pStrSymbol);
            AcSpec acSpecs;
            if (acSpecSymbol.Equals(TckFreqVar, StringComparison.OrdinalIgnoreCase))
            {
                acSpecs = WriteOneRecordToAcSpecForTck(acSpecSymbol, pStrValue, pStrTyp, pStrMin, pStrMax);
            }
            else if (index >= 0)
            {
                acSpecs = WriteOneRecordToAcSpecForNwire(acSpecSymbol, pStrValue, pStrTyp, pStrMin, pStrMax, index);
            }
            else
            {
                acSpecs = WriteOneRecordToAcSpec(acSpecSymbol, pStrValue, pStrTyp, pStrMin, pStrMax);
            }

            _acSpecSheet.AddRow(acSpecs);
        }

        internal AcSpec WriteOneRecordToAcSpec(string pStrSymbol, string pStrValue, string pStrTyp, string pStrMin, string pStrMax)
        {
            //Write basic data
            var acSpecs = new AcSpec(pStrSymbol, GetSelectorList(), pStrValue);
            //Write Category
            foreach (string categroyName in _acFullCategoryList)
            {
                var categroy = new CategoryInSpec(categroyName, pStrTyp, pStrMin, pStrMax);
                acSpecs.AddCategory(categroy);
            }
            return acSpecs;
        }

        internal AcSpec WriteOneRecordToAcSpecForTck(string pStrSymbol, string pStrValue, string pStrTyp, string pStrMin, string pStrMax)
        {
            //Write basic data
            var acSpecs = new AcSpec(pStrSymbol, GetSelectorList(), pStrValue);
            //Write Category
            foreach (string categroyName in _acFullCategoryList)
            {
                if (categroyName.Equals(SpiRomCat, StringComparison.OrdinalIgnoreCase))
                {
                    //Spi_rom change the frequency to 10MHZ
                    string value = 10e6.ToString(CultureInfo.InvariantCulture);
                    var categroy = new CategoryInSpec(categroyName, value, value, value);
                    acSpecs.AddCategory(categroy);
                }
                else
                {
                    var categroy = new CategoryInSpec(categroyName, pStrTyp, pStrMin, pStrMax);
                    acSpecs.AddCategory(categroy);
                }
            }
            return acSpecs;
        }

        internal AcSpec WriteOneRecordToAcSpecForNwire(string pStrSymbol, string pStrValue, string pStrTyp, string pStrMin, string pStrMax, int pinIdex)
        {
            DataTable table = NwireSingleton.Instance().SettingInfo.SettingTable;

            //Write basic data
            var acSpecs = new AcSpec(pStrSymbol, GetSelectorList(), pStrValue);
            //Write Category
            int actionRow = 1;
            foreach (string categroyName in _acFullCategoryList)
            {
                if (categroyName.Contains("Default"))
                {
                    string flowControl = table.Rows[actionRow][0].ToString();
                    string controloActoiin = table.Rows[actionRow][pinIdex + 1].ToString();
                    if (Regex.IsMatch(controloActoiin, "Enable@", RegexOptions.IgnoreCase) &&
                        Regex.IsMatch(categroyName, flowControl, RegexOptions.IgnoreCase))
                    {
                        string frequency = controloActoiin;
                        string valueFreq = Regex.Match(frequency, TestPlanConst.UnitRegPattern).Groups[TestPlanConst.Value].ToString();
                        string unit = Regex.Match(frequency, TestPlanConst.UnitRegPattern).Groups[TestPlanConst.Unit].ToString();

                        if (valueFreq.TryCombineHz(unit, out string targetFreq))
                        {
                            var categroy = new CategoryInSpec(categroyName, targetFreq, targetFreq, targetFreq);
                            acSpecs.AddCategory(categroy);
                        }
                    }
                    else
                    {
                        var categroy = new CategoryInSpec(categroyName, pStrTyp, pStrMin, pStrMax);
                        acSpecs.AddCategory(categroy);
                    }
                    actionRow++;
                }
                else
                {
                    var categroy = new CategoryInSpec(categroyName, pStrTyp, pStrMin, pStrMax);
                    acSpecs.AddCategory(categroy);
                }
            }
            return acSpecs;
        }

        private string GetAcSpecSymbol(string symbolName)
        {
            return SpecFormat.GenAcSpecSymbol(symbolName);
        }

        private void UpdateSymbolSpecs(Dictionary<string, double> tSetVarValueDict)
        {
            var lVarList = new List<string>();
            foreach (AcSpec data in _acSpecSheet.Rows)
            {
                lVarList.Add(data.Symbol.ToUpper());
            }
            foreach (KeyValuePair<string, double> timing in tSetVarValueDict)
            {
                if (lVarList.Contains(timing.Key.ToUpper()))
                {
                    continue;
                }
                var acSpecs = new AcSpec(timing.Key, GetSelectorList());
                List<string> acFullCategoryList = _acSpecSheet.CategoryList;
                // Write Category
                foreach (string categroyName in acFullCategoryList)
                {
                    const string value = BasicInitial.AcSpecDefault;
                    var categroy = new CategoryInSpec(categroyName, value, value, value);
                    acSpecs.AddCategory(categroy);
                }
                _acSpecSheet.AddRow(acSpecs);
            }
        }

        private string TSetMapCategory(Dictionary<string, double> tSetVarValueDict, string tSetBlock)
        {
            var regexblock = new Regex(tSetBlock, RegexOptions.IgnoreCase | RegexOptions.Compiled);
            List<string> targetCategorys = _acSpecSheet.CategoryList.FindAll(regexblock.IsMatch);
            var checkValueDict = new Dictionary<string, Dictionary<string, bool>>();
            foreach (string targetCategory in targetCategorys)
            {
                var subDict = new Dictionary<string, bool>();
                foreach (KeyValuePair<string, double> tSetItem in tSetVarValueDict)
                {
                    List<AcSpec> acSymbolDatas = _acSpecSheet.Rows.ToList().FindAll(a => a.Symbol.ToUpper().Equals(tSetItem.Key.ToUpper()));

                    foreach (AcSpec acSymbolData in acSymbolDatas)
                    {
                        if (acSymbolData.ContainsCategory(targetCategory))
                        {
                            CategoryInSpec categoryItem = acSymbolData.GetCategoryItem(targetCategory);
                            if (categoryItem.Min.Equals(tSetItem.Value.ToString()))
                            {
                                if (!subDict.ContainsKey(tSetItem.Key))
                                {
                                    subDict.Add(tSetItem.Key, true);
                                }
                                continue;
                            }
                            if (categoryItem.Min == BasicInitial.AcSpecDefault) // update value
                            {
                                categoryItem.Min = tSetItem.Value.ToString();
                                categoryItem.Typ = tSetItem.Value.ToString();
                                categoryItem.Max = tSetItem.Value.ToString();
                                if (!subDict.ContainsKey(tSetItem.Key))
                                {
                                    subDict.Add(tSetItem.Key, true);
                                }
                                continue;
                            }
                            if (categoryItem.Name.Equals("SPI_ROM") && !tSetItem.Value.Equals(-1))
                            {
                                categoryItem.Min = tSetItem.Value.ToString();
                                categoryItem.Typ = tSetItem.Value.ToString();
                                categoryItem.Max = tSetItem.Value.ToString();
                                if (!subDict.ContainsKey(tSetItem.Key))
                                {
                                    subDict.Add(tSetItem.Key, true);
                                }
                                continue;
                            }


                            if (!categoryItem.Min.Equals(tSetItem.Value.ToString()))
                            {
                                if (!subDict.ContainsKey(tSetItem.Key))
                                {
                                    subDict.Add(tSetItem.Key, false);
                                }
                            }
                        }
                    }
                }
                checkValueDict.Add(targetCategory, subDict);
            }

            string returnCate = tSetBlock;

            foreach (KeyValuePair<string, Dictionary<string, bool>> checkcategoryItem in checkValueDict)
            {
                if (checkcategoryItem.Value.All(p => p.Value))  // all true no need create cateogy
                {
                    returnCate = checkcategoryItem.Key;
                    return returnCate;
                }
            }
            if (tSetBlock == BlockType.SPI_ROM.ToString())
            {
                return returnCate;
            }

            if (targetCategorys.Count == 0)
            {
                returnCate = tSetBlock;
            }
            else
            {
                returnCate = tSetBlock + "_" + targetCategorys.Count;
            }

            AddAcCategory(tSetVarValueDict, returnCate);

            return returnCate;
        }

        private void AddAcCategory(Dictionary<string, double> tSetVarValueDict, string categoryName)
        {
            if (_acSpecSheet.CategoryList.Contains(categoryName))
            {
                return;
            }

            _acSpecSheet.CategoryList.Add(categoryName);
            foreach (AcSpec specs in _acSpecSheet.Rows)
            {
                string value;
                if (!tSetVarValueDict.ContainsKey(specs.Symbol.ToUpper()))
                {
                    value = specs.CategoryList[0].Typ;
                }
                else
                {
                    value = tSetVarValueDict[specs.Symbol.ToUpper()].ToString();
                }
                var newCategory = new CategoryInSpec(categoryName, value, value, value);
                specs.AddCategory(newCategory);
            }
        }

        protected bool CopyAcCategory(string tarCategory, string newCategoryName)
        {
            if (_acSpecSheet.CategoryList.Contains(newCategoryName))
            {
                return false;
            }

            _acSpecSheet.CategoryList.Add(newCategoryName);
            int tarCateIndex = _acSpecSheet.CategoryList.FindIndex(a => a.Equals(tarCategory, StringComparison.OrdinalIgnoreCase));
            if (tarCateIndex == -1)
            {
                return false;
            }

            foreach (AcSpec specs in _acSpecSheet.Rows)
            {
                var newCategory = new CategoryInSpec(newCategoryName, specs.CategoryList[tarCateIndex].Typ,
                    specs.CategoryList[tarCateIndex].Min, specs.CategoryList[tarCateIndex].Max);
                specs.AddCategory(newCategory);
            }
            return true;
        }

        private List<Selector> GetSelectorList()
        {
            var selectorList = new List<Selector>
            {
                new Selector("Typ", "Typ"), new Selector("Min", "Min"), new Selector("Max", "Max")
            };
            return selectorList;
        }
        #endregion
    }
}
