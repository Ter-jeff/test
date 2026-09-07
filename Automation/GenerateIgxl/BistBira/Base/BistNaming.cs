using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Utility;

using ScghLib.Reader;

using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.BistBira.Base
{
    public class BistNaming
    {
        private readonly Dictionary<string, List<KeyAndPosition>> _repairFlags;
        private readonly Dictionary<string, MbistProductionSpecialNamingRule> _specialNamingRules;
        private static string _moduleCpu = "";
        private static string _moduleGpu = "";
        private static string _moduleSoc = "";

        public string ModuleCpu
        {
            get { return _moduleCpu; }
        }

        public string ModuleGpu
        {
            get { return _moduleGpu; }
        }

        public string ModuleSoc
        {
            get { return _moduleSoc; }
        }

        public BistNaming(MbistConfig pConfig)
        {
            _moduleCpu = ModuleSingleton.Instance().ModuleCpu;
            _moduleGpu = ModuleSingleton.Instance().ModuleGfx;
            _moduleSoc = ModuleSingleton.Instance().ModuleSoc;

            _repairFlags = pConfig.RepairFlagSetting;
            _specialNamingRules = pConfig.SpecialNamingRules;

        }

        public bool IsBira(string pattern, string label)
        {
            string[] pat = pattern.ToUpper().Split('_');

            if ((pat.Contains("PLLP") && (pat.Contains("BIR") || pat.Contains("BIRA"))) || pat.Contains("FLS"))
            {
                return true;
            }

            if (Regex.IsMatch(pattern, "(BIRA)|(_BIRA*_*)|(_FLS)", RegexOptions.IgnoreCase))
            {
                return true;
            }

            if (Regex.IsMatch(pattern, "_ERT", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(label, "_BIRA", RegexOptions.IgnoreCase))
            {
                return true;
            }

            if (Regex.IsMatch(pattern, "_RETENTION", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(pattern, "BIRA", RegexOptions.IgnoreCase))
            {
                return true;
            }

            return false;

        }

        public static string GetHarvestInstName(string sheetName)
        {
            return "Harvest_Summary_After_" + sheetName;
        }

        private string CreateTestNameFromMapping(string pModule, BistProdFlowRow bistProdFlowSheetRow, string fieldFromLabel, bool vMarginHasNumber)
        {
            var results = new List<string>();

            if (string.IsNullOrEmpty(bistProdFlowSheetRow.Pattern))
            {
                return "";
            }

            //ModuleBlock (ex: SocMbist)
            string moduleBlock = pModule + BistConst.ConMbist;
            results.Add(moduleBlock);

            //Multiple Mbist Sheet for Ellis
            if (bistProdFlowSheetRow.IsMultipleMbistScghSheet)
            {
                results.Add(bistProdFlowSheetRow.SheetName.Split('_').Last());
            }

            if (!string.IsNullOrEmpty(bistProdFlowSheetRow.TestName))
            {
                results.Add(bistProdFlowSheetRow.TestName);
                return Combination.CombineByUnderLine(results);
            }

            //P.Mode (ex: MC601)
            string performanceMode = GetPerformanceMode(bistProdFlowSheetRow, vMarginHasNumber);
            results.Add(performanceMode);

            if (bistProdFlowSheetRow.OtherPModeInfoStr != "")
            {
                results.Add(bistProdFlowSheetRow.OtherPModeInfoStr);
            }

            string keywordFromLabel = GetSubName(bistProdFlowSheetRow.Label, fieldFromLabel);
            if (keywordFromLabel != "")
            {
                results.Add(keywordFromLabel);
            }

            //dcCategory
            if (!Regex.IsMatch(bistProdFlowSheetRow.DcCategory, "error", RegexOptions.IgnoreCase))
            {
                string dcInfo = "";
                if (!string.IsNullOrEmpty(bistProdFlowSheetRow.EqnVoltage))
                {
                    dcInfo = bistProdFlowSheetRow.Voltage.Split(' ').First();
                }
                else if (!string.IsNullOrEmpty(bistProdFlowSheetRow.DcCategory))
                {
                    dcInfo = bistProdFlowSheetRow.DcCategory;
                }

                if (!dcInfo.Equals(performanceMode))
                {
                    results.Add(dcInfo);
                }
            }

            //STEP5. Payload
            string subNames = bistProdFlowSheetRow.Pattern;
            results.Add(subNames);

            //STEP5.1 EQN (Optional)
            if (bistProdFlowSheetRow.Voltage.Contains("_EQN") && LocalSpecs.EquationVoltagesFileName != "N/A")
            {
                TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.VbtFuncNameFunctionalTUpdated, "mbist");
                {
                    results.Add("EQN");
                }
            }
            //STEP6. HV/LV/NV
            string voltageType = bistProdFlowSheetRow.IsDsscRow ? GetVoltageType(bistProdFlowSheetRow.OriVoltage) : GetVoltageType(bistProdFlowSheetRow.Voltage);  //NV/LV/HV
            results.Add(voltageType);

            return Combination.CombineByUnderLine(results);
        }

        private static string GetPerformanceMode(BistProdFlowRow pRow, bool vMarginHasNumber)
        {
            string performanceMode;
            if (Regex.IsMatch(pRow.VoltageMode, "VMARGIN", RegexOptions.IgnoreCase))
            {
                performanceMode = !vMarginHasNumber
                    ? pRow.VoltageMode.Replace("1", "").Replace("3", "")
                    : pRow.VoltageMode;
            }
            else
            {
                performanceMode = PerformanceModeSingleton.Instance().FindPerformanceMode(pRow.VoltageMode);
                if (string.IsNullOrEmpty(performanceMode) && !string.IsNullOrEmpty(pRow.VoltageMode))
                {
                    performanceMode = pRow.VoltageMode.Split(',', ' ').Length > 1 ? pRow.VoltageMode.Split(',', ' ')[0] : pRow.VoltageMode;
                }
            }

            if (string.IsNullOrEmpty(performanceMode))
            {
                performanceMode = "ALLFRV";
            }
            return performanceMode;
        }

        public string CreateNewTestName(string pModule, BistProdFlowRow row)
        {
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                return $"SocMbist_{row.Pattern}_{row.Label.Split('_')[1]}_{row.Voltage.Split(',').Last()}";
            }

            string fieldsFromLabel = "";
            bool vMarginHasNumber = true;
            if (_specialNamingRules.ContainsKey(pModule.ToUpper()))
            {
                fieldsFromLabel = _specialNamingRules[pModule.ToUpper()].UseLabelPosition;
                vMarginHasNumber = _specialNamingRules[pModule.ToUpper()].VMarginNeedNumber;
            }
            return CreateTestNameFromMapping(pModule, row, fieldsFromLabel, vMarginHasNumber);
        }

        public string CreateDummyTestName(string pModule, BistProdFlowRow pRow)
        {
            var result = new List<string> { pModule + BistConst.ConMbist, "Dummy", pRow.Label };
            string voltageType = GetVoltageType(pRow.Voltage);  //NV/LV/HV
            result.Add(voltageType);

            return Combination.CombineByUnderLine(result);
        }

        public string GetModule(string sheetName)
        {
            if (Regex.IsMatch(sheetName, "^C.*", RegexOptions.IgnoreCase))
            {
                return _moduleCpu;
            }
            if (Regex.IsMatch(sheetName, "^G.*|^L.*", RegexOptions.IgnoreCase))
            {
                return _moduleGpu;
            }
            if (Regex.IsMatch(sheetName, "^S.*", RegexOptions.IgnoreCase))
            {
                return _moduleSoc;
            }
            return "";
        }

        public string GetPatternModule(string mode, string pCurrentModule)
        {
            string newModule = ModuleSingleton.GetModuleByPerformanceMode(mode);

            if (newModule != "" && newModule != pCurrentModule)
            {
                return newModule;
            }

            return pCurrentModule;
        }

        public string GetPatternType(string pPattern, bool mbistGet = false)
        {
            string subName = GetSubName(pPattern, 3);
            const string pattern = @"^IN\w*$|IN\d*$";
            if (Regex.IsMatch(subName, pattern, RegexOptions.IgnoreCase))
            {
                if (mbistGet)
                {
                    string sixSubName = GetSubName(pPattern, 6);
                    string eightSubName = GetSubName(pPattern, 8);
                    if (Regex.IsMatch(sixSubName, "(M2I)|(BIS)", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(eightSubName, "(FLS)|(EFC)", RegexOptions.IgnoreCase))
                    {
                        return BistConst.ConPayload;
                    }
                }
                return BistConst.ConInit;
            }
            return BistConst.ConPayload;
        }

        public bool IsSelDssc(string pPattern)
        {
            string subName = GetSubName(pPattern, 3);
            string pattern = @"^IN\w*$|IN\d*$";
            string patDssc = "SRMDSSC"; // jadecide DSSC_E
            return Regex.IsMatch(subName, pattern, RegexOptions.IgnoreCase) && Regex.IsMatch(pPattern, patDssc, RegexOptions.IgnoreCase);
        }

        public static string GetVoltageType(string pVoltage)
        {
            string type;
            string selector = "";
            if (pVoltage.Split(' ', ',').Length > 1)
            {
                string voltage = pVoltage.Split(' ', ',').First();
                selector = pVoltage.Split(' ', ',').Last();
                if (Regex.IsMatch(voltage, BistConst.ConVmargin1, RegexOptions.IgnoreCase))
                {
                    return BistConst.ConMhv;
                }

                if (Regex.IsMatch(voltage, BistConst.ConVmargin3, RegexOptions.IgnoreCase))
                {
                    return BistConst.ConMlv;
                }
            }
            if (Regex.IsMatch(selector, BistConst.ConVmargin1, RegexOptions.IgnoreCase) || Regex.IsMatch(selector, BistConst.ConMhv, RegexOptions.IgnoreCase))
            {
                if (Regex.IsMatch(selector, BistConst.ConMhv1, RegexOptions.IgnoreCase))
                {
                    type = BistConst.ConMhv1;
                }
                else if (Regex.IsMatch(selector, BistConst.ConMhv2, RegexOptions.IgnoreCase))
                {
                    type = BistConst.ConMhv2;
                }
                else if (Regex.IsMatch(selector, BistConst.ConMhv3, RegexOptions.IgnoreCase))
                {
                    type = BistConst.ConMhv3;
                }
                else
                {
                    type = BistConst.ConMhv;
                }
            }
            else if (Regex.IsMatch(selector, BistConst.ConVmargin3, RegexOptions.IgnoreCase) || Regex.IsMatch(selector, BistConst.ConMlv, RegexOptions.IgnoreCase))
            {
                if (Regex.IsMatch(selector, BistConst.ConMlv1, RegexOptions.IgnoreCase))
                {
                    type = BistConst.ConMlv1;
                }
                else if (Regex.IsMatch(selector, BistConst.ConMlv2, RegexOptions.IgnoreCase))
                {
                    type = BistConst.ConMlv2;
                }
                else if (Regex.IsMatch(selector, BistConst.ConMlv3, RegexOptions.IgnoreCase))
                {
                    type = BistConst.ConMlv3;
                }
                else
                {
                    type = BistConst.ConMlv;
                }
            }
            else if (Regex.IsMatch(selector, BistConst.ConVmax, RegexOptions.IgnoreCase) || Regex.IsMatch(selector, BistConst.ConHv, RegexOptions.IgnoreCase))
            {
                type = BistConst.ConHv;
            }
            else if (Regex.IsMatch(selector, BistConst.ConVmin, RegexOptions.IgnoreCase) || Regex.IsMatch(selector, BistConst.ConLv, RegexOptions.IgnoreCase))
            {
                type = BistConst.ConLv;
            }
            else if (Regex.IsMatch(selector, BistConst.ConVnom, RegexOptions.IgnoreCase) || Regex.IsMatch(selector, BistConst.ConNv, RegexOptions.IgnoreCase))
            {
                type = BistConst.ConNv;
            }
            else if (Regex.IsMatch(selector, BistConst.ConVdst, RegexOptions.IgnoreCase) || Regex.IsMatch(selector, BistConst.ConVdisturb, RegexOptions.IgnoreCase))
            {
                type = BistConst.ConLv;
            }
            else
            {
                type = BistConst.ConNv;
            }
            return type;
        }

        public static bool IsNormalVoltage(string pVoltage)
        {
            if (string.Equals(pVoltage, BistConst.ConVmargin1, StringComparison.OrdinalIgnoreCase) || string.Equals(pVoltage, BistConst.ConMhv, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(pVoltage, BistConst.ConVmargin3, StringComparison.OrdinalIgnoreCase) || string.Equals(pVoltage, BistConst.ConMlv, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(pVoltage, BistConst.ConVmax, StringComparison.OrdinalIgnoreCase) || string.Equals(pVoltage, BistConst.ConHv, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(pVoltage, BistConst.ConVmin, StringComparison.OrdinalIgnoreCase) || string.Equals(pVoltage, BistConst.ConLv, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(pVoltage, BistConst.ConVnom, StringComparison.OrdinalIgnoreCase) || string.Equals(pVoltage, BistConst.ConNv, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(pVoltage, BistConst.ConVdst, StringComparison.OrdinalIgnoreCase) || string.Equals(pVoltage, BistConst.ConVdisturb, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(pVoltage, BistConst.ConMhv1, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(pVoltage, BistConst.ConMhv2, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(pVoltage, BistConst.ConMhv3, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(pVoltage, BistConst.ConMlv1, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(pVoltage, BistConst.ConMlv2, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(pVoltage, BistConst.ConMlv3, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return false;
        }

        public BistBinTableType GetBinType(string pVoltage)
        {
            string levels = GetVoltageType(pVoltage);
            switch (levels)
            {
                case BistConst.ConMhv:
                    return BistBinTableType.BinMhv;
                case BistConst.ConMhv1:
                    return BistBinTableType.BinMhv1;
                case BistConst.ConMhv2:
                    return BistBinTableType.BinMhv2;
                case BistConst.ConMhv3:
                    return BistBinTableType.BinMhv3;
                case BistConst.ConMlv:
                    return BistBinTableType.BinMlv;
                case BistConst.ConMlv1:
                    return BistBinTableType.BinMlv1;
                case BistConst.ConMlv2:
                    return BistBinTableType.BinMlv2;
                case BistConst.ConMlv3:
                    return BistBinTableType.BinMlv3;
                case BistConst.ConLv:
                    return BistBinTableType.BinLv;
                case BistConst.ConHv:
                    return BistBinTableType.BinHv;
                case BistConst.ConNv:
                    return BistBinTableType.BinNv;
                default:
                    return BistBinTableType.BinNv;
            }
        }

        public string GetSubName(string name, string rule)
        {
            var resultList = new List<string>();
            string[] words = name.Split('_');
            string[] numbersStrings = rule.Split(',');
            foreach (string numbers in numbersStrings)
            {
                if (rule.ToLower() == "full")
                {
                    return name;
                }

                if (rule == "")
                {
                    //no operation
                }
                else
                {
                    int getNumber = int.Parse(numbers);
                    if (words.Length > getNumber && getNumber >= 0)
                    {
                        //resultName = resultName + "_" + words[getNumber];
                        resultList.Add(words[getNumber]);
                    }
                    //Error message
                }
            }
            string resultName = Combination.CombineByUnderLine(resultList);
            return resultName;
        }

        public string GetSubName(string pPattern, int pNumber)
        {
            string result = "";
            string[] subNameList = pPattern.Split('_');
            if (pNumber > 0 && pNumber < subNameList.Length)
            {
                result = subNameList[pNumber];
            }
            return result;
        }

        public bool JudgePattern(string pattern, string compareItem)
        {
            if (GetSubName(pattern, 8).Equals(compareItem, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        public bool JudgeRepairPattern(string pattern, string module)
        {
            bool match = false;
            List<KeyAndPosition> rules = _repairFlags.First(p => p.Key.Equals(module, StringComparison.OrdinalIgnoreCase)).Value;
            foreach (KeyAndPosition keyAndPosition in rules)
            {
                bool onceMatch = true;
                string[] positions = keyAndPosition.KeyPositions.Split(',');
                string[] keys = keyAndPosition.Keys.Split(',');
                if (!positions.Length.Equals(keys.Length))
                {
                    throw new Exception($"The RepairFlag setting key and keyposition for {module} is not matched!");
                }
                for (int i = 0; i < positions.Length; i++)
                {
                    if (!keys[i].Equals(GetSubName(pattern, positions[i]), StringComparison.OrdinalIgnoreCase))
                    {
                        onceMatch = false;
                    }
                }

                if (onceMatch)
                {
                    match = true;
                }
            }
            return match;
        }

        public string GetWaitTimeOnly(bool isWaitTimeOnly)
        {
            if (isWaitTimeOnly)
            {
                return "WaitTimeOnly";
            }

            return "";
        }

        public string CreateRetentionTestNameNew(string module, string category, string sheetName = "", string waitTime = "", string step = "", string isWaitTimeOnly = "")
        {
            var result = new List<string>
            {
                module + BistConst.ConMbist,
                sheetName,
                category,
            };
            if (waitTime != "")
            {
                result.Add("Wait" + waitTime + "mS");
            }

            if (step != "")
            {
                result.Add("STEP" + step);
            }

            result.Add(isWaitTimeOnly);
            result.RemoveAll(x => x.Equals(""));
            return string.Join("_", result);
        }
    }
}
