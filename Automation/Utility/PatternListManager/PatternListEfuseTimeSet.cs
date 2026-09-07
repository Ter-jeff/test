using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;

using TestPlanLib.Basic;

namespace Automation.Utility.PatternListManager
{
    public class PatternListEfuseTimeSet
    {
        private const string UsedType = "USE";

        private static VpnSettingFile _settingFile;
        private static Dictionary<string, string> _patternDigi4TypeDic;

        public static List<string> GetEfusePatternTimeSet(List<PatternData> patlists)
        {
            var efuseTimeSetSheetList = new List<string>();
            ReadCfg();
            foreach (PatternData pattern in patlists)
            {
                if (JudgeEfusePattern(pattern.PatternName))
                {
                    if (pattern.Use.ToUpper().Equals(UsedType))
                    {
                        if (pattern.FileVersion.Equals("NA", StringComparison.OrdinalIgnoreCase) ||
                            pattern.TimeSetVersion.Equals("NA", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }
                        if (!efuseTimeSetSheetList.Contains(pattern.TimeSetVersion))
                        {
                            efuseTimeSetSheetList.Add(pattern.TimeSetVersion);
                        }
                    }
                }
            }

            return efuseTimeSetSheetList;
        }

        public static void ReadCfg()
        {
            string infoConfig = Path.Combine(AppContext.BaseDirectory, "Config", "VpnSettingFile.xml");
            var reader = new System.Xml.Serialization.XmlSerializer(typeof(VpnSettingFile));
            var cfgXml = new StreamReader(infoConfig);
            _settingFile = (VpnSettingFile)reader.Deserialize(cfgXml);
            _settingFile.DicOrgStr =
                _settingFile.DicOrgStr.Replace("\t", "").Replace("\n", "").Replace("\r", "").Replace(" ", "").Replace("\'", "").Replace("\"", "");
            _settingFile.DicSpecStr =
                _settingFile.DicSpecStr.Replace("\t", "").Replace("\n", "").Replace("\r", "").Replace(" ", "").Replace("\'", "").Replace("\"", "");

            cfgXml.Close();

            _patternDigi4TypeDic = new Dictionary<string, string>();

            List<string> digi4List = _settingFile.DicSpecStr.Split(',').ToList();

            foreach (string s in digi4List)
            {
                string regexPattern = @"(?<key>\w+)\s*[:]\s*(?<value>\w+)";
                string key = Regex.Match(s, regexPattern).Groups["key"].ToString();
                string value = Regex.Match(s, regexPattern).Groups["value"].ToString();
                _patternDigi4TypeDic.Add(key, value);
            }
        }

        private static bool JudgeEfusePattern(string pattern)
        {
            const string efuse = "EFUSE";
            List<string> subName = pattern.Split('_').ToList();
            if (subName.Count > 5)
            {
                if (_patternDigi4TypeDic.ContainsKey(subName[4].ToUpper()) &&
                    _patternDigi4TypeDic[subName[4].ToUpper()] == efuse)
                {
                    return true;
                }
                if (subName.Count > 12 &&
                    _patternDigi4TypeDic[efuse.ToUpper()] == efuse)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
