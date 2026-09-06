using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using IgxlLib.IgxlSheets;

using TestPlanLib.Xml;

namespace Automation.Utility.Basic
{
    public static class DifferentialService
    {
        private static DiffPairConfig _config;

        private static DiffPairConfig Config
        {
            get
            {
                if (_config == null)
                {
                    if (File.Exists(Path.Combine(AppContext.BaseDirectory, "Config", "DiffPairConfig.xml")))
                    {

                        _config = XmlService<DiffPairConfig>.LoadXml(Path.Combine(AppContext.BaseDirectory, "Config", "DiffPairConfig.xml"));
                    }
                }
                return _config;
            }
        }

        public static Dictionary<string, string> DifferentialPair(List<string> pinList)
        {
            var pairs = new Dictionary<string, string>();
            foreach (DiffItem pinPair in Config.DiffPairPins)
            {
                //Add differential pairs from config which defined using Pin name
                string posPin = pinList.Find(p => p.Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase));
                string negPin = pinList.Find(p => p.Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase));
                if (posPin != null && negPin != null && !pairs.ContainsKey(posPin))
                {
                    pairs.Add(posPin, negPin);
                }
            }

            for (int i = 0; i < pinList.Count; i++)
            {
                for (int j = i + 1; j < pinList.Count; j++)
                {
                    if (pinList[i].Length == pinList[j].Length)
                    {
                        GetSamePartInDiffPairs(pinList[i], pinList[j], out string nStr, out string pStr);
                        bool flag = false;
                        foreach (DiffItem rule in Config.DiffPairRules)
                        {
                            if (Regex.IsMatch(pStr, rule.Pos, RegexOptions.IgnoreCase) && nStr.Equals(Regex.Replace(pStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.ContainsKey(pinList[j]))
                                {
                                    pairs.Add(pinList[j], pinList[i]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                            else if (Regex.IsMatch(nStr, rule.Pos, RegexOptions.IgnoreCase) && pStr.Equals(Regex.Replace(nStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.ContainsKey(pinList[i]))
                                {
                                    pairs.Add(pinList[i], pinList[j]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                        }
                        if (flag)
                        {
                            break;
                        }
                    }
                }
            }

            //Store Neg and Pos pin again
            int count = pairs.Count;
            for (int i = 0; i < count; i++)
            {
                pairs.Add(pairs.ElementAt(i).Value, pairs.ElementAt(i).Key);
            }
            return pairs;
        }

        public static List<string> GroupDiffPairs(List<string> oriPinList)
        {
            var pinList = oriPinList.ToList();
            DiffPairConfig config = XmlService<DiffPairConfig>.LoadXml(Path.Combine(AppContext.BaseDirectory, "Config", "DiffPairConfig.xml"));
            var pairs = new List<string>();

            foreach (DiffItem pinPair in Config.DiffPairPins)
            {
                //Add differential pairs from config which defined using Pin name
                string posPin = pinList.Find(p => p.Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase));
                string negPin = pinList.Find(p => p.Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase));
                if (posPin != null && negPin != null && !pairs.Contains(posPin + "::" + negPin))
                {
                    pairs.Add(posPin + "::" + negPin);
                    pinList.Remove(posPin);
                    pinList.Remove(negPin);
                }
            }

            for (int i = 0; i < pinList.Count; i++)
            {
                for (int j = i + 1; j < pinList.Count; j++)
                {
                    if (pinList[i].Length == pinList[j].Length)
                    {
                        GetSamePartInDiffPairs(pinList[i], pinList[j], out string nStr, out string pStr);
                        bool flag = false;
                        foreach (DiffItem rule in config.DiffPairRules)
                        {
                            if (Regex.IsMatch(pStr, rule.Pos, RegexOptions.IgnoreCase) &&
                                nStr.Equals(Regex.Replace(pStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.Contains(pinList[j] + "::" + pinList[i]))
                                {
                                    pairs.Add(pinList[j] + "::" + pinList[i]);
                                    pinList.Remove(pinList[j]);
                                    pinList.Remove(pinList[i]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                            else if (Regex.IsMatch(nStr, rule.Pos, RegexOptions.IgnoreCase) &&
                                pStr.Equals(Regex.Replace(nStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")), StringComparison.OrdinalIgnoreCase))
                            {
                                if (!pairs.Contains(pinList[i] + "::" + pinList[j]))
                                {
                                    pairs.Add(pinList[i] + "::" + pinList[j]);
                                    pinList.Remove(pinList[j]);
                                    pinList.Remove(pinList[i]);
                                    i--;
                                    flag = true;
                                    break;
                                }
                            }
                        }
                        if (flag)
                        {
                            break;
                        }
                    }
                }
            }

            pairs.AddRange(pinList);
            return pairs;
        }

        public static bool DiffPinPosAndNeg(string diffPins, out string pos, out string neg, out string groupName)
        {
            pos = "";
            neg = "";
            groupName = "";
            bool isDiff = false;
            if (!diffPins.Contains("::"))
            {
                return false;
            }

            string[] pair = diffPins.Split(new[] { "::" }, StringSplitOptions.None);
            for (int i = 0; i < pair.Length; i++)
            {
                pair[i] = pair[i].Trim();
            }

            #region Diff group name by rule-1 in DiffPairPins
            foreach (DiffItem pinPair in Config.DiffPairPins)
            {
                if (pair[0].Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase)
                    && pair[1].Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase))
                {
                    pos = pair[0];
                    neg = pair[1];
                    groupName = pos;
                    isDiff = true;
                }
                if (pair[1].Equals(pinPair.Pos, StringComparison.OrdinalIgnoreCase)
                        && pair[0].Equals(pinPair.Neg, StringComparison.OrdinalIgnoreCase))
                {
                    pos = pair[1];
                    neg = pair[0];
                    groupName = pos;
                    isDiff = true;
                }
            }
            #endregion

            #region Diff group name by rule-2 get common string and compare diff. string

            GetSamePartInDiffPairs(ref groupName, pair[0], pair[1], out string nStr, out string pStr);

            foreach (DiffItem rule in Config.DiffPairRules)
            {
                if (Regex.IsMatch(pStr, rule.Pos, RegexOptions.IgnoreCase) &&
                    nStr.Equals(Regex.Replace(pStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")),
                        StringComparison.OrdinalIgnoreCase))
                {
                    pos = pair[0];
                    neg = pair[1];
                    isDiff = true;
                }
                else if (Regex.IsMatch(nStr, rule.Pos, RegexOptions.IgnoreCase) &&
                         pStr.Equals(Regex.Replace(nStr, rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")),
                             StringComparison.OrdinalIgnoreCase))
                {
                    pos = pair[1];
                    neg = pair[0];
                    isDiff = true;
                }
            }
            #endregion

            #region Diff group name by rule-3
            //EX:ADDR_M2P_DQ_N::ADDR_M2P_DQ_P
            foreach (DiffItem rule in Config.DiffPairRules)
            {
                if (Regex.IsMatch(pair[0], rule.Pos, RegexOptions.IgnoreCase) &&
                    pair[1].Equals(Regex.Replace(pair[0], rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")),
                        StringComparison.OrdinalIgnoreCase))
                {
                    string newGroupName = Regex.Replace(pair[0], rule.Pos, "");
                    groupName = newGroupName.Length < groupName.Length ? newGroupName : groupName;
                    pos = pair[0];
                    neg = pair[1];
                    isDiff = true;
                }
                else if (Regex.IsMatch(pair[1], rule.Pos, RegexOptions.IgnoreCase) &&
                         pair[0].Equals(Regex.Replace(pair[1], rule.Pos, rule.Neg.Replace("^", "").Replace("$", "")),
                             StringComparison.OrdinalIgnoreCase))
                {
                    string newGroupName = Regex.Replace(pair[1], rule.Pos, "");
                    groupName = newGroupName.Length < groupName.Length ? newGroupName : groupName;
                    pos = pair[1];
                    neg = pair[0];
                    isDiff = true;
                }
            }
            #endregion

            if (isDiff)
            {
                if (!string.IsNullOrEmpty(groupName))
                {
                    groupName += groupName.EndsWith("_", StringComparison.OrdinalIgnoreCase) ? "Diff" : "_Diff";
                }
            }
            else
            {
                groupName = "";
            }

            return isDiff;
        }

        private static void GetSamePartInDiffPairs(ref string groupName, string nPinName, string pPinName, out string nStr, out string pStr)
        {
            //EX:ADDR_M2P_DQ_N0::ADDR_M2P_DQ_P0
            pStr = "";
            nStr = "";
            if (nPinName.Length == pPinName.Length)
            {
                for (int i = 0; i < nPinName.Length; i++)
                {
                    if (nPinName[i] == pPinName[i])
                    {
                        groupName += nPinName[i];
                    }
                    else
                    {
                        pStr += nPinName[i];
                        nStr += pPinName[i];
                    }
                }
            }
        }

        private static void GetSamePartInDiffPairs(string nPinName, string pPinName, out string nStr, out string pStr)
        {
            //EX:ADDR_M2P_DQ_N0::ADDR_M2P_DQ_P0
            pStr = "";
            nStr = "";
            if (nPinName.Length == pPinName.Length)
            {
                for (int i = 0; i < nPinName.Length; i++)
                {
                    if (nPinName[i] != pPinName[i])
                    {
                        nStr += nPinName[i];
                        pStr += pPinName[i];
                    }
                }
            }
        }

        public static string GenDiffGroupName(PinMapSheet pinMapSheet, string diffPinName, bool isNeedGenPinGroup)
        {
            string groupName = "";
            if (!diffPinName.Contains("::"))
            {
                return diffPinName;
            }

            if (!isNeedGenPinGroup)
            {
                return string.Join(",", Regex.Split(diffPinName, "::"));
            }
            if (pinMapSheet != null)
            {
                string[] pair = diffPinName.Split(new[] { "::" }, StringSplitOptions.None);
                groupName = pinMapSheet.GetDiffGroupName(pair);
            }

            if (groupName == "")
            {
                DiffPinPosAndNeg(diffPinName, out _, out _, out groupName);
            }
            return groupName == "" ? diffPinName : groupName;
        }
    }
}
