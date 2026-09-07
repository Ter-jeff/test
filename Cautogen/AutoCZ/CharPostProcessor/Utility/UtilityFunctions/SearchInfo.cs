using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

using IgxlLib.IgxlBase;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions
{
    public class SearchInfo
    {
        /* Member function */
        public static HardIpReference GetHardIpInfo(string patternName)
        {
            string pattern = patternName.Split(':')[0].ToLower();
            if (LocalSpecs.PatInfoList == null)
            {
                return null;
            }

            return LocalSpecs.PatInfoList.TryGetValue(pattern, out HardIpReference value)
                ? value
                : null;
        }

        public static bool CheckPatUsed(CharPlanItem planItem)
        {
            return string.IsNullOrEmpty(planItem.IsNeedMask);
            //return (_IsUsedPat(planItem.InitPattern1)
            //        && _IsUsedPat(planItem.InitPattern2)
            //        && _IsUsedPat(planItem.InitPattern3)
            //        && _IsUsedPat(planItem.InitPattern4)
            //        && _IsUsedPat(planItem.InitPattern5)
            //        && _IsUsedPat(planItem.InitPattern6)
            //        && _IsUsedPat(planItem.InitPattern7)
            //        && _IsUsedPat(planItem.InitPattern8)
            //        && _IsUsedPat(planItem.InitPattern9)
            //        && _IsUsedPat(planItem.InitPattern10)
            //        && _IsUsedPat(planItem.Payload1)
            //        && _IsUsedPat(planItem.Payload2)
            //        && _IsUsedPat(planItem.Payload3)
            //        && _IsUsedPat(planItem.Payload4)
            //        && _IsUsedPat(planItem.Payload5));
            //return !planItem.UsedPatterns.Any(x => _IsUsedPat(x) == false);
        }

        public static List<string> GetExtendLimits(string testName, string status)
        {
            var testNameList = new List<string>();
            string[] userDefList = testName.Split('_');
            string measType = userDefList[1];
            string measPin = userDefList[5];

            if (DataConvertor.CheckMixedPins(status, measType) && LocalSpecs.ProgInfo.PinGroupDic.TryGetValue(measPin, out PinGroup value))
            {
                List<Pin> pins = value.PinList;
                foreach (Pin pin in pins.OrderBy(p => p.PinName))
                {
                    userDefList[5] = pin.PinName.Replace("_", "");
                    testNameList.Add(string.Join("_", userDefList));
                }
            }
            else
            {
                testNameList.Add(testName);
            }

            return testNameList;
        }

        public static List<string> GetLimitsOfPinFromPinGroup(string testName, bool genCSharp)
        {
            var testNameList = new List<string>();
            string[] userDefList = testName.Split('_');
            string measType = userDefList[1];
            string measPin = userDefList[5];

            if (LocalSpecs.ProgInfo.PinGroupDic.TryGetValue(measPin, out PinGroup value))
            {
                List<Pin> pins = value.PinList;
                foreach (Pin pin in pins.OrderBy(p => p.PinName))
                {
                    if (genCSharp)
                    {
                        testNameList.Add(pin.PinName);
                    }
                    else
                    {
                        userDefList[5] = pin.PinName.Replace("_", "");
                        testNameList.Add(string.Join("_", userDefList));
                    }
                }
            }
            else
            {
                if (genCSharp)
                {
                    testNameList.Add(LocalSpecs.ProgInfo.PinDic.ContainsKey(measPin) ? LocalSpecs.ProgInfo.PinDic[measPin] : measPin);
                }
                else
                {
                    testNameList.Add(testName);
                }
            }

            return testNameList;
        }

        public static string GetTimeset(string patternName)
        {
            string pattern = patternName.Split(':')[0].ToUpper();
            if (!LocalSpecs.PatternDatas.ContainsKey(pattern))
            {
                return "";
            }

            return Regex.IsMatch(LocalSpecs.PatternDatas[pattern].TimesetVersion, "^timeset_", RegexOptions.IgnoreCase)
                ? LocalSpecs.PatternDatas[pattern].TimesetVersion
                : "";
        }

        public static bool IsContains(string main, string pattern)
        {
            if (main.ToUpper().Contains(pattern.ToUpper()))
            {
                return true;
            }

            return false;
        }

    }
}
