using System.Collections.Generic;

using CommonLib.Extension;

using TestPlanLib.DataStruct;

namespace TestPlanLib.Utility
{
    public static class MultiTestSettingUtility
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
        public const string ValtRowPinNameFlag = "_Valt";
        public const string MsgNotFoundDefaultCategory = "Can not find default category for: '{0}' in testSettings!";
        public const string MsgNotFoundMbistEfuseCategory = "Can not find Mbist Efuse Category for: '{0}' in testSetting!";
        public const string MsgNotFoundMbistRetentionCategory = "Can not find Mbist Retention Category for: '{0}','{1}' in testSetting!";
        public const string MsgNotFoundRtosCategory = "Can not find Rtos category for performanceMode: '{0}' in testSettings!";
        public const string MsgNotFoundHardIpDcCategory = "Can not find HardIpDc Category for: '{0}' in testSetting!";
        public const string MsgNotFoundPmDcCategory = "Can not find performanceMode: '{0}' for '{1}' in testSettings!";
        public const string MsgNotFoundDcCategory = "Can not find '{0}' Dc category for: domain '{1}',subtest '{2}',performanceMode '{3}', pattern '{4}' in testSettings";
        public const string MsgDuplicateDcCategory = "Duplicate '{0}' Dc category for: domain '{1}',subtest '{2}',performanceMode '{3}', pattern '{4}' in testSettings";
        public const string MsgNotFoundBinCutDcCategory = "Can not find BinCut default Dc Category!";

        public static bool ExistChipletVddPin(string pinName, string chiplet, PowerInfoSheet powerInfoSheet)
        {
            PowerInfoSheet powerInfo = powerInfoSheet;
            if (powerInfo != null)
            {
                foreach (PowerInfoRow row in powerInfo.Rows)
                {
                    if (row.Chiplet.EqualsIgnoreCase(chiplet) && row.PinName.EqualsIgnoreCase(pinName))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public static Dictionary<int, List<DcCategoryInfo>> CompareUserDefine(List<DcCategoryInfo> dcCategoryInfos, List<string> patterns, string chiplet)
        {
            var dicUserdefineMatchedCount = new Dictionary<int, List<DcCategoryInfo>>();
            if (patterns == null || patterns.Count == 0)
            {
                return dicUserdefineMatchedCount;
            }

            string userDefKey = string.IsNullOrEmpty(chiplet) ? DcCategoryName.CategoryDefaultValue : chiplet;

            var patternBlocklst = new List<string>();
            if (patterns[0].Contains('_'))
            {
                patternBlocklst = patterns != null ? [.. patterns[0].Split('_')] : [];
            }

            foreach (DcCategoryInfo category in dcCategoryInfos)
            {
                //if category has no userDefine block , set match count = 0
                if (category.UserDefined.EqualsIgnoreCase(userDefKey))
                {
                    if (!dicUserdefineMatchedCount.TryGetValue(0, out List<DcCategoryInfo>? value))
                    {
                        value = [];
                        dicUserdefineMatchedCount.Add(0, value);
                    }

                    value.Add(category);
                    continue;
                }

                //Compare all the userDefine blocks
                List<string> categoryUserDefinelst = [.. category.UserDefined.Split('_')];
                int matchedCount = 0;
                foreach (string userdefine in categoryUserDefinelst)
                {
                    if (patternBlocklst.Exists(s => s.EqualsIgnoreCase(userdefine)))
                    {
                        matchedCount++;
                    }
                    else
                    {
                        matchedCount = -1;
                        break;
                    }
                }
                if (matchedCount > -1)
                {
                    if (!dicUserdefineMatchedCount.TryGetValue(matchedCount, out List<DcCategoryInfo>? value))
                    {
                        value = [];
                        dicUserdefineMatchedCount.Add(matchedCount, value);
                    }

                    value.Add(category);
                }
            }
            return dicUserdefineMatchedCount;
        }
    }
}
