using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Enums;
using CommonLib.Extension;
using CommonLib.Utility;

using TestPlanLib.Const;
using TestPlanLib.Utility;

namespace TestPlanLib.DataStruct
{
    public partial class DcCategoryInfos : List<DcCategoryInfo>
    {
        public const string HardIp = "HardIP";
        public const string Ids = "IDS";
        public const string Mbist = "Mbist";
        public const string Efuse = "Efuse";
        public const string Evs = "EVS";
        public const string Bincut = "BINCUT";
        public const string Sram = "Sram";
        public const string Logic = "Logic";
        public const string Stress = "Stress";
        public const string Retention = "Retention";
        public const string Bist = "Bist";
        public const string Bira = "Bira";
        public const string Rtos = "Rtos";
        public const string Nwire = "Nwire";
        public const string Conti = "Conti";
        public const string Td = "Td";
        public const string Tdchain = "TdChain";
        public const string Sa = "Sa";
        public const string Sachain = "SaChain";
        public const string Init = "Init";
        public const string Boot = "BOOT";

        public const string MsgNotFoundDefaultCategory = "Can not find default category for: '{0}' in testSettings!";
        public const string MsgNotFoundMbistEfuseCategory = "Can not find Mbist Efuse Category for: '{0}' in testSetting!";
        public const string MsgNotFoundMbistRetentionCategory = "Can not find Mbist Retention Category for: '{0}','{1}' in testSetting!";
        public const string MsgNotFoundRtosCategory = "Can not find Rtos category for performanceMode: '{0}' in testSettings!";
        public const string MsgNotFoundHardIpDcCategory = "Can not find HardIpDc Category for: '{0}' in testSetting!";
        public const string MsgNotFoundPmDcCategory = "Can not find performanceMode: '{0}' for '{1}' in testSettings!";
        public const string MsgNotFoundDcCategory = "Can not find '{0}' Dc category for: domain '{1}',subtest '{2}',performanceMode '{3}', pattern '{4}' in testSettings";
        public const string MsgDuplicateDcCategory = "Duplicate '{0}' Dc category for: domain '{1}',subtest '{2}',performanceMode '{3}', pattern '{4}' in testSettings";
        public const string MsgNotFoundBinCutDcCategory = "Can not find BinCut default Dc Category!";
        private const string ErrorShouldhave = " (Error,ShouldHave)";

        [GeneratedRegex("^(?!Mbist)M([a-zA-Z]){1}([a-zA-Z0-9]){2}([a-zA-Z0-9]){1,2}$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex("VRS", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        private static readonly Regex _regex1 = MyRegex();
        private static readonly Regex _regex = MyRegex1();

        public string ContainsCategoryName(List<string> keys)
        {
            if (Exists(s => keys.All(y => s.CategoryName.Split('_').ToList().Contains(y, StringExtensions.IgnoreCase))))
            {
                IEnumerable<string> arr = this.Where(s => keys.All(y => s.CategoryName.Split('_').ToList().Contains(y, StringExtensions.IgnoreCase))).Select(x => x.CategoryName);
                return string.Join(",", arr);
            }
            return "";
        }

        public bool IsSplitDcSpecs(bool optionalsIsSplitDcSpecs)
        {
            bool isSplitDcSpecs = optionalsIsSplitDcSpecs;
            return Count > 500 || isSplitDcSpecs;
        }

        public List<string> GetChipletList(PowerInfoSheet powerInfoSheet)
        {
            var chipletList = new List<string>();
            var chipletcategory = this.Select(x => x.Chiplet(powerInfoSheet)).Distinct().ToList();
            PowerInfoSheet powerInfo = powerInfoSheet;
            if (powerInfo != null)
            {
                foreach (string s in powerInfo.Rows.Select(x => x.Chiplet).Distinct().ToList())
                {
                    if (chipletcategory.Exists(x => x.EqualsIgnoreCase(s)))
                    {
                        chipletList.Add(s);
                    }
                }
                if (chipletList.Count == 0)
                {
                    chipletList.Add("");
                }
            }

            return chipletList;
        }

        /// <summary>
        /// Mbist_[domain]_Init_VRS or Mbist_[domain]_Init_X
        /// </summary>
        /// <returns></returns>
        public string FindMbistVrsInitCategory(string domain, string chiplet = "")
        {
            string userDefKey = string.IsNullOrEmpty(chiplet) ? DcCategoryName.CategoryDefaultValue : chiplet;
            DcCategoryInfo? mbistVrsInitCategory = Find(s => s.Test.EqualsIgnoreCase(Mbist) && s.Domain.EqualsIgnoreCase(domain) && s.Subtest.EqualsIgnoreCase(Init) && s.PmodePatternVdip.EqualsIgnoreCase("VRS") && s.UserDefined.EqualsIgnoreCase(userDefKey)) ?? Find(s => s.Test.EqualsIgnoreCase(Mbist) && s.Domain.EqualsIgnoreCase(domain) && s.Subtest.EqualsIgnoreCase(Init) && s.PmodePatternVdip.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) && s.UserDefined.EqualsIgnoreCase(userDefKey));
            return mbistVrsInitCategory == null ? "" : mbistVrsInitCategory.CategoryName;
        }

        public string FindMbistEfuseCategory(string domain, string chiplet)
        {
            string userDefKey = string.IsNullOrEmpty(chiplet) ? DcCategoryName.CategoryDefaultValue : chiplet;
            DcCategoryInfo? mbistEfuseDcCategory = Find(s => !s.IsHardipDcCategory &&
                                                                             s.Test.EqualsIgnoreCase(Mbist) &&
                                                                             s.Domain.EqualsIgnoreCase(domain) &&
                                                                             s.Subtest.EqualsIgnoreCase(Efuse) &&
                                                                             s.PmodePatternVdip.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                             s.UserDefined.EqualsIgnoreCase(userDefKey));
            return mbistEfuseDcCategory == null ? FindSpecialDefaultCategory(HardIp) : mbistEfuseDcCategory.CategoryName;
        }

        ///  <summary>
        /// Eg: Nwire_X_X_X, Conti_X_X_X, IDS_X_X_X, HardIP_X_X_X, RTOS_X_X_X, BINCUT_X_X_X
        ///  </summary>
        ///  <param name="blockName"></param>
        /// <param name="chiplet"></param>
        /// <returns></returns>
        public string FindSpecialDefaultCategory(string blockName, string chiplet = "")
        {
            string userDefKey = string.IsNullOrEmpty(chiplet) ? DcCategoryName.CategoryDefaultValue : chiplet;
            DcCategoryInfo? defaultCategory = Find(s => s.Test.EqualsIgnoreCase(blockName) &&
                                                                        s.Domain.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                        s.Subtest.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                        s.PmodePatternVdip.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                        s.UserDefined.EqualsIgnoreCase(userDefKey));

            return defaultCategory == null ? "" : defaultCategory.CategoryName;
        }

        /// <summary>
        /// Find HardIpDc DcCategory Name
        /// </summary>
        /// <param name="hardipDcItemName"></param>
        /// <returns></returns>
        public string FindHardIpDcCategory(string hardipDcItemName)
        {
            DcCategoryInfo? hardipDcCategory = Find(s => s.Test.EqualsIgnoreCase(HardIp) &&
                                                                         s.Domain.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                         s.Subtest.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                         s.PmodePatternVdip.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                         s.UserDefined.EqualsIgnoreCase(hardipDcItemName));
            return hardipDcCategory == null ? "" : hardipDcCategory.CategoryName;
        }

        public string FindDcCategoryName(string test, PowerInfoSheet powerInfoSheet, string domain = DcCategoryName.CategoryDefaultValue, string subtest = DcCategoryName.CategoryDefaultValue, string performanceMode = DcCategoryName.CategoryDefaultValue, string chiplet = "")
        {
            DcCategoryInfo? categoryInfo = Find(s => s.Test.EqualsIgnoreCase(test) &&
                                                                     s.Domain.EqualsIgnoreCase(domain) &&
                                                                     s.Subtest.EqualsIgnoreCase(subtest) &&
                                                                     s.PmodePatternVdip.EqualsIgnoreCase(performanceMode) &&
                                                                     s.Chiplet(powerInfoSheet).EqualsIgnoreCase(chiplet));
            return categoryInfo != null ? categoryInfo.CategoryName : "";
        }

        /// <summary>
        /// Find Nwire category "Nwire_X_X_X"
        /// </summary>
        /// <param name="enumMessageLevel"></param>
        /// <param name="errorMsg"></param>
        /// <returns></returns>
        public string FindNwireCategory(out EnumMessageLevel enumMessageLevel, out string errorMsg)
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";

            string nwireDefaultCategory = FindSpecialDefaultCategory(Nwire);
            if (string.IsNullOrEmpty(nwireDefaultCategory))
            {
                enumMessageLevel = EnumMessageLevel.Error;
                errorMsg = string.Format(MsgNotFoundDefaultCategory, Nwire);
            }
            return nwireDefaultCategory;
        }

        /// <summary>
        /// Find Conti category "Conti_X_X_X"
        /// </summary>
        /// <param name="enumMessageLevel"></param>
        /// <param name="errorMsg"></param>
        /// <returns></returns>
        public string FindContiCategory(out EnumMessageLevel enumMessageLevel, out string errorMsg, string chiplet = "")
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";

            string contiDefaultCategory = FindSpecialDefaultCategory(Conti, chiplet);
            if (string.IsNullOrEmpty(contiDefaultCategory))
            {
                enumMessageLevel = EnumMessageLevel.Error;
                errorMsg = string.Format(MsgNotFoundDefaultCategory, Conti);
            }
            return contiDefaultCategory;
        }

        public string FindRtosCategory(PowerInfoSheet powerInfoSheet, string performanceMode, out EnumMessageLevel enumMessageLevel, out string errorMsg)
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";

            string rtosDefaultCategory = FindDcCategoryName(Rtos, powerInfoSheet, DcCategoryName.CategoryDefaultValue, Boot);
            if (string.IsNullOrEmpty(performanceMode))
            {
                if (string.IsNullOrEmpty(rtosDefaultCategory))
                {
                    enumMessageLevel = EnumMessageLevel.Error;
                    errorMsg = string.Format(MsgNotFoundDefaultCategory, Rtos);
                }
                return rtosDefaultCategory;
            }

            DcCategoryInfo? userDefineRtosCatgory = Find(s => !s.IsHardipDcCategory &&
                                                                              s.Test.EqualsIgnoreCase(Rtos) &&
                                                                              s.Domain.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                              s.Subtest.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue) &&
                                                                              s.PmodePatternVdip.EqualsIgnoreCase(performanceMode) &&
                                                                              s.UserDefined.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue));
            if (userDefineRtosCatgory == null)
            {
                enumMessageLevel = EnumMessageLevel.Error;
                errorMsg = string.Format(MsgNotFoundRtosCategory, performanceMode);
                return "";
            }
            return userDefineRtosCatgory.CategoryName;
        }

        public string FindMbistCatgeoryName(PowerInfoSheet powerInfoSheet, string sheetDomain, string type, string performanceMode, List<string> patterns, out EnumMessageLevel enumMessageLevel, out string errorMsg, string chiplet = "", string pmodeDomain = "")
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";
            bool searchFailurePmodeDomain = false;
            string dcCategoryByPmodeDomain = "";
            bool searchFailureSheetDomain = false;
            string dcCategoryBySheetDomain = "";

            if (!string.IsNullOrEmpty(pmodeDomain) && pmodeDomain != sheetDomain)
            {
                dcCategoryByPmodeDomain = FindMbistCatgeoryNameByOneModule(powerInfoSheet, pmodeDomain, type, performanceMode, patterns, out enumMessageLevel, out errorMsg, out searchFailurePmodeDomain, chiplet);
            }

            if (!string.IsNullOrEmpty(sheetDomain))
            {
                dcCategoryBySheetDomain = FindMbistCatgeoryNameByOneModule(powerInfoSheet, sheetDomain, type, performanceMode, patterns, out enumMessageLevel, out errorMsg, out searchFailureSheetDomain, chiplet);
            }

            if (string.IsNullOrEmpty(dcCategoryByPmodeDomain))
            {
                return dcCategoryBySheetDomain;
            }

            if (searchFailurePmodeDomain && searchFailureSheetDomain)
            {
                return dcCategoryByPmodeDomain;
            }

            return !searchFailurePmodeDomain ? dcCategoryByPmodeDomain : dcCategoryBySheetDomain;
        }

        public string FindMbistCatgeoryNameByOneModule(PowerInfoSheet powerInfoSheet, string domain, string type, string performanceMode, List<string> patternlst, out EnumMessageLevel enumMessageLevel, out string errorMsg, out bool searchFailure, string chiplet = "")
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";
            string dcCategory = "";
            searchFailure = false;
            performanceMode = string.IsNullOrEmpty(performanceMode) ? DcCategoryName.CategoryDefaultValue : performanceMode;
            type = string.IsNullOrEmpty(type) ? DcCategoryName.CategoryDefaultValue : type;
            domain = string.IsNullOrEmpty(domain) ? DcCategoryName.CategoryDefaultValue : domain;

            //Efuse
            if (type.EqualsIgnoreCase(Efuse))
            {
                dcCategory = FindMbistEfuseCategory(domain, chiplet);
                if (string.IsNullOrEmpty(dcCategory))
                {
                    enumMessageLevel = EnumMessageLevel.Error;
                    errorMsg = string.Format(MsgNotFoundMbistEfuseCategory, domain);
                    dcCategory = Combination.CombineByUnderLine("Mbist_" + domain + "_Efuse_X", chiplet) + ErrorShouldhave;
                }
                return dcCategory;
            }

            //Retention
            if (performanceMode.EqualsIgnoreCase(Retention))
            {
                dcCategory = FindMbistRetentionCategory(powerInfoSheet, domain, type, chiplet);
                if (string.IsNullOrEmpty(dcCategory))
                {
                    enumMessageLevel = EnumMessageLevel.Error;
                    errorMsg = string.Format(MsgNotFoundMbistRetentionCategory, domain, type);
                    dcCategory = Combination.CombineByUnderLine("Mbist_" + domain + "_" + type + "_" + Retention, chiplet) + ErrorShouldhave;
                }
                return dcCategory;
            }

            //Retention pattern category
            if (type.EqualsIgnoreCase(MBistConst.ConErtbira) ||
                type.EqualsIgnoreCase(MBistConst.ConErtbist) ||
                type.EqualsIgnoreCase(MBistConst.ConNrt) ||
                type.EqualsIgnoreCase(MBistConst.ConWus) ||
                type.EqualsIgnoreCase(MBistConst.ConSrt) ||
                type.EqualsIgnoreCase(MBistConst.ConEfc))
            {
                dcCategory = FindMbistRetionPatternCategory(powerInfoSheet, domain, type, performanceMode, chiplet);
                if (string.IsNullOrEmpty(dcCategory))
                {
                    enumMessageLevel = EnumMessageLevel.Error;
                    errorMsg = string.Format(MsgNotFoundMbistRetentionCategory, domain, type);
                    dcCategory = Combination.CombineByUnderLine("Mbist_" + domain + "_" + type + "_" + performanceMode, chiplet) + ErrorShouldhave;
                }
                return dcCategory;
            }

            //Find valid categorys(type=[Bist/Bira] or type="X")
            List<DcCategoryInfo> validCategorylst = GetValidCategorylst(powerInfoSheet, domain, type, performanceMode, chiplet);

            //Compare UserDefine
            Dictionary<int, List<DcCategoryInfo>> dicUserdefineMatchedCount = MultiTestSettingUtility.CompareUserDefine(validCategorylst, patternlst, chiplet);

            if (validCategorylst.Count == 0 || dicUserdefineMatchedCount.Keys.Count == 0)
            {
                //if mode is VRS, use Mbist_[domain]_Init_VRS
                if (_regex.IsMatch(performanceMode))
                {
                    dcCategory = FindMbistVrsInitCategory(domain, chiplet);
                }
                //if np performanceMode, use Init category
                else if (string.IsNullOrEmpty(performanceMode) || performanceMode == DcCategoryName.CategoryDefaultValue)
                {
                    dcCategory = FindMbistVrsInitCategory(domain, chiplet);
                }

                if (string.IsNullOrEmpty(dcCategory))
                {
                    enumMessageLevel = EnumMessageLevel.Error;
                    string pattern = patternlst != null ? string.Join(",", patternlst) : "";
                    errorMsg = string.Format(MsgNotFoundDcCategory, Mbist, domain, type, performanceMode, pattern);

                    searchFailure = true;
                    dcCategory = Combination.CombineByUnderLine("Mbist_" + domain + "_" + type + "_" + performanceMode, chiplet) + ErrorShouldhave;
                }
                return dcCategory;
            }

            return GetCategoryByUserdefine(domain, type, performanceMode, patternlst, ref enumMessageLevel, ref errorMsg, ref dicUserdefineMatchedCount);
        }

        public static string GetCategoryByUserdefine(string domain, string type, string performanceMode, List<string> patternlst, ref EnumMessageLevel enumMessageLevel, ref string errorMsg, ref Dictionary<int, List<DcCategoryInfo>> dicUserdefineMatchedCount)
        {
            if (type.EqualsIgnoreCase(Bist) || type.EqualsIgnoreCase(Bira))
            {
                var dicBistBiraUserdefineMatchedCount = new Dictionary<int, List<DcCategoryInfo>>();

                //If type=[Bira/Bist], select [Bira/Bist] categorys
                foreach (int key in dicUserdefineMatchedCount.Keys)
                {
                    List<DcCategoryInfo> bistbiraCategorys = dicUserdefineMatchedCount[key].FindAll(s => s.Subtest.EqualsIgnoreCase(type));
                    if (bistbiraCategorys.Count > 0)
                    {
                        dicBistBiraUserdefineMatchedCount.Add(key, bistbiraCategorys);
                    }
                }
                //If type=[Bist], select [Bist] categorys, if no [Bist], use [BIRA] first
                if (dicBistBiraUserdefineMatchedCount.Count == 0 && type.EqualsIgnoreCase(Bist))
                {
                    foreach (int key in dicUserdefineMatchedCount.Keys)
                    {
                        List<DcCategoryInfo> bistbiraCategorys = dicUserdefineMatchedCount[key].FindAll(s => s.Subtest.EqualsIgnoreCase(Bira));
                        if (bistbiraCategorys.Count > 0)
                        {
                            dicBistBiraUserdefineMatchedCount.Add(key, bistbiraCategorys);
                        }
                    }
                }
                //If can not find [Bira/Bist] categorys, use X
                if (dicBistBiraUserdefineMatchedCount.Count == 0)
                {
                    foreach (int key in dicUserdefineMatchedCount.Keys)
                    {
                        List<DcCategoryInfo> bistbiraCategorys = dicUserdefineMatchedCount[key].FindAll(s => s.Subtest.EqualsIgnoreCase(DcCategoryName.CategoryDefaultValue));
                        if (bistbiraCategorys.Count > 0)
                        {
                            dicBistBiraUserdefineMatchedCount.Add(key, bistbiraCategorys);
                        }
                    }
                }

                dicUserdefineMatchedCount = dicBistBiraUserdefineMatchedCount;
            }

            //find the max match count and use the first one category
            int maxMatchedCount = dicUserdefineMatchedCount.Keys.Max();
            if (dicUserdefineMatchedCount[maxMatchedCount].Count > 1)
            {
                //Warnning:More than one dcCategory found in testSettings
                enumMessageLevel = EnumMessageLevel.Warning;
                string pattern = patternlst != null ? string.Join(",", patternlst) : "";
                errorMsg = string.Format(MsgDuplicateDcCategory, Mbist, domain, type, performanceMode, pattern);
            }
            return string.Join(",", dicUserdefineMatchedCount[maxMatchedCount].Select(s => s.CategoryName));
        }

        public List<DcCategoryInfo> GetValidCategorylst(PowerInfoSheet powerInfoSheet, string domain, string type, string performanceMode, string chiplet)
        {
            List<DcCategoryInfo> validCategorylst = FindDcCategoryByKeys(powerInfoSheet, Mbist, domain, type, performanceMode, chiplet);
            if (type.EqualsIgnoreCase(Bist))
            {
                validCategorylst.AddRange(FindDcCategoryByKeys(powerInfoSheet, Mbist, domain, Bira, performanceMode, chiplet));
            }
            if (type.EqualsIgnoreCase(Bira) || type.EqualsIgnoreCase(Bist))
            {
                validCategorylst.AddRange(FindDcCategoryByKeys(powerInfoSheet, Mbist, domain, DcCategoryName.CategoryDefaultValue, performanceMode, chiplet));
            }

            return validCategorylst;
        }

        public string FindHardIpCatgeoryName(PowerInfoSheet powerInfoSheet, string performanceMod, string hardipDcItemName, string hardIpBlockName, List<string> patternlst, string chiplet, out EnumMessageLevel enumMessageLevel, out string errorMsg)
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";
            string dcCategory;

            //If performanceMod is not a valid  performanceMode(pattern not contain performanceMode), set performanceMod = "X"
            if (!_regex1.IsMatch(performanceMod))
            {
                performanceMod = DcCategoryName.CategoryDefaultValue;
            }

            string hardipDefaultCategory = FindSpecialDefaultCategory(HardIp, chiplet);
            //HardIpDc category: if cannot find it, flag error
            if (!string.IsNullOrEmpty(hardipDcItemName))
            {
                dcCategory = FindHardIpDcCategory(hardipDcItemName);
                if (string.IsNullOrEmpty(dcCategory))
                {
                    enumMessageLevel = EnumMessageLevel.Error;
                    errorMsg = string.Format(MsgNotFoundHardIpDcCategory, hardipDcItemName);
                    return "";
                }
                return dcCategory;
            }

            //Find all the hardip categorys by performanceMod
            List<DcCategoryInfo> validCategorylst = FindDcCategoryByKeys(powerInfoSheet, HardIp, hardIpBlockName, DcCategoryName.CategoryDefaultValue, performanceMod, chiplet);
            Dictionary<int, List<DcCategoryInfo>> dicUserdefineMatchedCount = MultiTestSettingUtility.CompareUserDefine(validCategorylst, patternlst, chiplet);

            if ((validCategorylst.Count == 0 || dicUserdefineMatchedCount.Count == 0) && !string.IsNullOrEmpty(hardIpBlockName))
            {
                validCategorylst = FindDcCategoryByKeys(powerInfoSheet, HardIp, DcCategoryName.CategoryDefaultValue, DcCategoryName.CategoryDefaultValue, performanceMod, chiplet);
                dicUserdefineMatchedCount = MultiTestSettingUtility.CompareUserDefine(validCategorylst, patternlst, chiplet);
            }

            if (validCategorylst.Count == 0)
            {
                if (performanceMod != DcCategoryName.CategoryDefaultValue)
                {
                    //Flag error and use hardIp default category: Can not find HardIp PerformanceMode category :{0} in testSettings!
                    enumMessageLevel = EnumMessageLevel.Error;
                    errorMsg = string.Format(MsgNotFoundPmDcCategory, performanceMod, HardIp);
                }
                return hardipDefaultCategory;
            }

            if (dicUserdefineMatchedCount.Keys.Count == 0)
            {
                if (string.IsNullOrEmpty(chiplet))
                {
                    return hardipDefaultCategory;
                }

                //Flag error:Exist Dc Category for performanceMode " ", but all the category userdefine not match pattern "".
                enumMessageLevel = EnumMessageLevel.Error;
                string pattern = patternlst != null ? string.Join(",", patternlst) : "";
                errorMsg = string.Format(MsgNotFoundDcCategory, HardIp, "", "", performanceMod, pattern);
                return "";
            }

            //find the max match count and use the first one category
            int maxMatchedCount = dicUserdefineMatchedCount.Keys.Max();
            if (dicUserdefineMatchedCount[maxMatchedCount].Count > 1)
            {
                //Warnning:More than one dcCategory found in testSettings
                enumMessageLevel = EnumMessageLevel.Warning;
                string pattern = patternlst != null ? string.Join(",", patternlst) : "";
                errorMsg = string.Format(MsgDuplicateDcCategory, HardIp, "", "", performanceMod, pattern);
            }
            return string.Join(",", dicUserdefineMatchedCount[maxMatchedCount].Select(s => s.CategoryName));
        }

        public string FindIdsCatgeoryName(string performanceMod, string hardipDcItemName, List<string> patternlst, out EnumMessageLevel enumMessageLevel, out string errorMsg)
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";
            string dcCategory;

            string idsDefaultCategory = FindSpecialDefaultCategory(Ids);

            //HardIpDc category: if cannot find it, flag error
            if (!string.IsNullOrEmpty(hardipDcItemName))
            {
                dcCategory = FindHardIpDcCategory(hardipDcItemName);
                if (string.IsNullOrEmpty(dcCategory))
                {
                    enumMessageLevel = EnumMessageLevel.Error;
                    errorMsg = string.Format(MsgNotFoundHardIpDcCategory, hardipDcItemName);
                    return "";
                }
                return dcCategory;
            }

            //Change by laura at 2019/09/19 For IDS,just have two Category: "IDS_X_X_X" OR "IDS_X_X_Off", if pattern contains _Off, use "IDS_X_X_Off" else use ""IDS_X_X_X"
            if (patternlst?.Count > 0)
            {
                if (patternlst.Exists(s => s.Contains("_OFF", StringComparison.OrdinalIgnoreCase)))
                {
                    DcCategoryInfo? category = Find(s => s.CategoryName.EqualsIgnoreCase("IDS_X_X_Off"));
                    if (category != null)
                    {
                        return category.CategoryName;
                    }
                }
            }
            //use IDS_X_X_X
            if (string.IsNullOrEmpty(idsDefaultCategory))
            {
                enumMessageLevel = EnumMessageLevel.Error;
                errorMsg = string.Format(MsgNotFoundDefaultCategory, Ids);
            }
            return idsDefaultCategory;
        }

        public string FindScanCategoryName(PowerInfoSheet powerInfoSheet, string test, string sheetDomain, string performanceMode, List<string> patterns, out EnumMessageLevel enumMessageLevel, out string errorMsg, string pmodeDomain = "", string chiplet = "")
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";
            string dcCategory = "";
            if (!string.IsNullOrEmpty(sheetDomain))
            {
                dcCategory = FindScanCategoryNameByOneModule(powerInfoSheet, test, sheetDomain, performanceMode, patterns, out enumMessageLevel, out errorMsg, chiplet);
            }

            if (string.IsNullOrEmpty(dcCategory) && pmodeDomain != sheetDomain)
            {
                dcCategory = FindScanCategoryNameByOneModule(powerInfoSheet, test, pmodeDomain, performanceMode, patterns, out enumMessageLevel, out errorMsg, chiplet);
            }

            return dcCategory;
        }

        public string FindScanCategoryNameByOneModule(PowerInfoSheet powerInfoSheet, string test, string domain, string performanceMode, List<string> patterns, out EnumMessageLevel enumMessageLevel, out string errorMsg, string chiplet = "")
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";

            performanceMode = string.IsNullOrEmpty(performanceMode) ? DcCategoryName.CategoryDefaultValue : performanceMode;
            domain = string.IsNullOrEmpty(domain) ? DcCategoryName.CategoryDefaultValue : domain;

            //SA, SaChain has no performanceMode
            if (test.EqualsIgnoreCase(Sa) || test.EqualsIgnoreCase(Sachain))
            {
                performanceMode = DcCategoryName.CategoryDefaultValue;
            }

            List<DcCategoryInfo> validCategorylst = FindDcCategoryByKeys(powerInfoSheet, test, domain, DcCategoryName.CategoryDefaultValue, performanceMode, chiplet);
            Dictionary<int, List<DcCategoryInfo>> dicUserdefineMatchedCount = MultiTestSettingUtility.CompareUserDefine(validCategorylst, patterns, chiplet);

            //if SaChain and TdChain , will find SaChain and TdChain first, if can not find any pattern, will find Sa/Td instead
            if (validCategorylst.Count == 0 || dicUserdefineMatchedCount.Keys.Count == 0)
            {
                if (test.EqualsIgnoreCase(Sachain) || test.EqualsIgnoreCase(Tdchain))
                {
                    validCategorylst = FindDcCategoryByKeys(powerInfoSheet, test[..^5], domain, DcCategoryName.CategoryDefaultValue, performanceMode, chiplet);
                    dicUserdefineMatchedCount = MultiTestSettingUtility.CompareUserDefine(validCategorylst, patterns, chiplet);
                }
            }

            if (validCategorylst.Count == 0 || dicUserdefineMatchedCount.Keys.Count == 0)
            {
                //if performanceMode exist in mbist, use mbist category
                if (performanceMode != DcCategoryName.CategoryDefaultValue)
                {
                    string dcCategory = FindDcCategoryName(Mbist, powerInfoSheet, domain, DcCategoryName.CategoryDefaultValue, performanceMode, chiplet);
                    if (!string.IsNullOrEmpty(dcCategory))
                    {
                        return dcCategory;
                    }
                }
                string dcCategory2 = FindDcCategoryName(test, powerInfoSheet, domain, DcCategoryName.CategoryDefaultValue, performanceMode, chiplet);
                if (!string.IsNullOrEmpty(dcCategory2))
                {
                    return dcCategory2;
                }

                //flag error
                enumMessageLevel = EnumMessageLevel.Error;
                string pattern = patterns != null ? string.Join(",", patterns) : "";
                errorMsg = string.Format(MsgNotFoundDcCategory, test, domain, "", performanceMode, pattern);
                return "";
            }

            //find the max match count and use the first one category
            int maxMatchedCount = dicUserdefineMatchedCount.Keys.Max();
            if (dicUserdefineMatchedCount[maxMatchedCount].Count > 1)
            {
                //Warnning:More than one dcCategory found in testSettings
                enumMessageLevel = EnumMessageLevel.Warning;
                string pattern = patterns != null ? string.Join(",", patterns) : "";
                errorMsg = string.Format(MsgDuplicateDcCategory, test, domain, "", performanceMode, pattern);
            }
            return string.Join(",", dicUserdefineMatchedCount[maxMatchedCount].Select(s => s.CategoryName));
        }

        public string FindBinCutCategoryName(PowerInfoSheet powerInfoSheet, string userDefine, string domain, out EnumMessageLevel enumMessageLevel, out string errorMsg)
        {
            enumMessageLevel = EnumMessageLevel.Result;
            errorMsg = "";

            if (!string.IsNullOrEmpty(userDefine))
            {
                string dcCategory = FindDcCategoryName(Bincut, powerInfoSheet, DcCategoryName.CategoryDefaultValue, DcCategoryName.CategoryDefaultValue, DcCategoryName.CategoryDefaultValue);
                if (!string.IsNullOrEmpty(dcCategory))
                {
                    return dcCategory;
                }
            }

            string binCutDefaultCategory = FindSpecialDefaultCategory(Bincut);
            if (string.IsNullOrEmpty(binCutDefaultCategory))
            {
                enumMessageLevel = EnumMessageLevel.Error;
                errorMsg = string.Format(MsgNotFoundBinCutDcCategory);
                return "BinCut_X_X_X";
            }

            return binCutDefaultCategory;
        }

        public List<DcCategoryInfo> FindDcCategoryByKeys(PowerInfoSheet powerInfoSheet, string test, string domain, string type, string performanceMode, string chiplet = "")
        {
            performanceMode = string.IsNullOrEmpty(performanceMode) ? DcCategoryName.CategoryDefaultValue : performanceMode;
            type = string.IsNullOrEmpty(type) ? DcCategoryName.CategoryDefaultValue : type;
            domain = string.IsNullOrEmpty(domain) ? DcCategoryName.CategoryDefaultValue : domain;

            return FindAll(s => !s.IsHardipDcCategory &&
                                                 s.Test.EqualsIgnoreCase(test) &&
                                                 s.Domain.EqualsIgnoreCase(domain) &&
                                                 s.Subtest.EqualsIgnoreCase(type) &&
                                                 s.PmodePatternVdip.EqualsIgnoreCase(performanceMode) &&
                                                 s.Chiplet(powerInfoSheet).EqualsIgnoreCase(chiplet));
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="domain"></param>
        /// <param name="type">ERTBIRA,ERTBIST,NRT,SRT,WUS</param>
        /// <returns></returns>
        public string FindMbistRetionPatternCategory(PowerInfoSheet powerInfoSheet, string domain, string type, string performanceMode, string chiplet)
        {
            string mbistDcCategory = FindDcCategoryName(Mbist, powerInfoSheet, domain, type, performanceMode, chiplet);
            if (string.IsNullOrEmpty(mbistDcCategory) && performanceMode != DcCategoryName.CategoryDefaultValue)
            {
                mbistDcCategory = FindDcCategoryName(Mbist, powerInfoSheet, domain, type, DcCategoryName.CategoryDefaultValue, chiplet);
            }
            if (string.IsNullOrEmpty(mbistDcCategory) && type.EqualsIgnoreCase(MBistConst.ConErtbist))
            {
                mbistDcCategory = FindDcCategoryName(Mbist, powerInfoSheet, domain, MBistConst.ConErtbira, performanceMode, chiplet);
                if (string.IsNullOrEmpty(mbistDcCategory) && performanceMode != DcCategoryName.CategoryDefaultValue)
                {
                    mbistDcCategory = FindDcCategoryName(Mbist, powerInfoSheet, domain, DcCategoryName.CategoryDefaultValue, chiplet);
                }
            }
            return mbistDcCategory;
        }

        public string FindMbistRetentionCategory(PowerInfoSheet powerInfoSheet, string domain, string type, string chiplet)
        {
            string mbistRetentionDcCategory = FindDcCategoryName(Mbist, powerInfoSheet, domain, type, Retention, chiplet);
            if (string.IsNullOrEmpty(mbistRetentionDcCategory) && type.EqualsIgnoreCase(MBistConst.ConErtbist))
            {
                mbistRetentionDcCategory = FindDcCategoryName(Mbist, powerInfoSheet, domain, MBistConst.ConErtbira, Retention, chiplet);
            }
            return mbistRetentionDcCategory;
        }
    }
}
