using System;
using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Static;

using CommonLib.Extension;

namespace BinCutScriptLib.Base.Line
{
    public class BvLine : BinCutLineBase
    {
        public BvLineInfo GetBvLineInfo()
        {
            //BV_VDD_SOC_MS001,2,VDD_CPU=0.543,VDD_GPU=0.596,VDD_SOC=0.606,VDD_CPU_SRAM=0.750,VDD_GPU_SRAM=0.668,VDD_FIXED=0.843,VDD_LOW=0.746
            var bvLineInfo = new BvLineInfo();
            string[] spt = Line.Split([','], StringSplitOptions.RemoveEmptyEntries);
            if (spt.Length > 2)
            {
                bvLineInfo.BvName = spt[0];

                if (!int.TryParse(spt[1], out bvLineInfo.Site))
                {
                    throw new Exception($"Site error @ {LineNo}: {Line}");
                }
            }
            string type = Reg.RegexType.Match(Line).Groups["type"].ToString();
            string[] types = type.Trim().Split(',');
            bvLineInfo.Type = types.First().Trim();
            bvLineInfo.InstType = type.Length >= 2 ? types[1].Trim() : "";
            bvLineInfo.LineOnly = Line.Replace("(" + type + ")", "").Trim().TrimEnd(',');
            bvLineInfo.Line = this;
            if (Line.Contains("EQN ="))
            {
                bvLineInfo.Eqn = int.Parse(spt[2].Replace("EQN =", "").Trim());
                bvLineInfo.Passbin = int.Parse(spt[3].Replace("PASSBIN =", "").Trim());
                bvLineInfo.C = double.Parse(spt[4].Replace("C =", "").Trim());
                bvLineInfo.M = double.Parse(spt[5].Replace("M =", "").Trim());
                bvLineInfo.Cpgb = double.Parse(spt[6].Replace("CPGB =", "").Replace("GB =", "").Trim());
                bvLineInfo.Ids = double.Parse(spt[7].Replace("IDS =", "").Replace("mA", "").Trim());
            }
            return bvLineInfo;
        }

        public BvLineInfo GetBvLineInfoCsharp()
        {
            var bvLineInfo = new BvLineInfo();
            if (Line.Contains("Bincut payload voltage"))
            {
                int site = GetSite();
                bvLineInfo.Site = site;
                bvLineInfo.LineOnly = Line;
                bvLineInfo.Type = "payload voltage";
                bvLineInfo.Line = this;
            }
            else if (Line.Contains("Bincut safe voltage"))
            {
                int site = GetSite();
                bvLineInfo.Site = site;
                bvLineInfo.LineOnly = Line;
                bvLineInfo.Type = "safe voltage";
                bvLineInfo.Line = this;
            }
            return bvLineInfo;
        }

        public BvLineInfo GetBvLineInfoCombinePayloadCs()
        {
            var bvLineInfo = new BvLineInfo();
            if (Line.Contains("Bincut payload voltage"))
            {
                int site = GetSite();
                bvLineInfo.Site = site;
                bvLineInfo.LineOnly = Line;
                bvLineInfo.Type = "payload voltage";
                bvLineInfo.Line = this;
            }
            return bvLineInfo;
        }

        public Dictionary<string, double> GetValues()
        {
            string[] spt1 = Line.Split([','], StringSplitOptions.RemoveEmptyEntries);
            var dic = new Dictionary<string, double>();
            for (int i = 2; i < spt1.Length; i++)
            {
                if (spt1[i].Contains('='))
                {
                    (string name, double value) = GetPair('=', spt1, i);
                    dic.Add(name, value);
                }
            }
            return dic;
        }

        public Dictionary<string, double> GetValuesCsharp()
        {
            string voltages = Line.Split(["Bincut payload voltage: "], StringSplitOptions.None).Last();
            string[] spt1 = voltages.Split([','], StringSplitOptions.RemoveEmptyEntries);
            var dic = new Dictionary<string, double>();
            for (int i = 0; i < spt1.Length; i++)
            {
                if (spt1[i].Contains(':'))
                {
                    (string name, double value) = GetPair(':', spt1, i);
                    dic.Add(name, value);
                }
            }
            return dic;
        }

        private static (string name, double value) GetPair(char key, string[] spt1, int i)
        {
            string[] arr = spt1[i].Split(key);
            string name = arr[0].Trim();
            _ = double.TryParse(arr[1].Replace("V", "").Replace("v", ""), out double value);
            return (name.ToUpper(), value);
        }

        public string GetText()
        {
            string[] spt1 = Line.Split([','], StringSplitOptions.RemoveEmptyEntries);
            var texts = new List<string>();
            for (int i = 2; i < spt1.Length; i++)
            {
                texts.Add(spt1[i]);
            }
            return string.Join(",", texts);
        }

        public Dictionary<string, double> GenDictionary()
        {
            return BvLineUtil.GenDictionary(() => Line);
        }
    }

    public class BvLineInfo
    {
        public string LineOnly = string.Empty;
        public int Site;
        public string BvName = string.Empty;
        public string Type = string.Empty;
        public string InstType = string.Empty;
        public BvLine Line = new();

        public int Eqn;
        public int Passbin;
        public double C;
        public double M;
        public double Cpgb;
        public double Ids;

        public bool IsHardip;

        public Dictionary<string, double> GenDictionary()
        {
            return BvLineUtil.GenDictionary(() => LineOnly);
        }

        public static bool CompareBVline(string lineA, string lineB)
        {
            string[] logSpt = lineA.Split([','], StringSplitOptions.RemoveEmptyEntries);
            string[] cmpSpt = lineB.Split([','], StringSplitOptions.RemoveEmptyEntries);

            if (logSpt.Length != cmpSpt.Length)
            {
                return false;
            }

            int compareLength = Math.Min(logSpt.Length, cmpSpt.Length);
            for (int i = 2; i < compareLength; i++)
            {
                int pos1 = logSpt[i].IndexOf('=');
                int pos2 = cmpSpt[i].IndexOf('=');

                string numStr1 = logSpt[i][(pos1 + 1)..];
                string numStr2 = cmpSpt[i][(pos2 + 1)..];

                if (double.TryParse(numStr1, out double logValue) &&
                    double.TryParse(numStr2, out double cmpValue))
                {
                    if (logValue != cmpValue)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public static bool CompareBVlineCs(string line1, string line2)
        {
            Dictionary<string, double> dict1 = ParseBvLineCs(line1);
            Dictionary<string, double> dict2 = ParseBvLineCs(line2);

            if (dict1.Count != dict2.Count)
            {
                return false;
            }

            foreach (KeyValuePair<string, double> kv in dict1)
            {
                if (!dict2.TryGetValue(kv.Key, out double v2))
                {
                    return false;
                }
                if (Math.Round(Math.Abs(kv.Value - v2), 3) > 0.001)
                {
                    return false;
                }
            }
            return true;
        }

        private static Dictionary<string, double> ParseBvLineCs(string line)
        {
            var dict = new Dictionary<string, double>(StringExtensions.IgnoreCase);
            foreach (string part in line.Split([','], StringSplitOptions.RemoveEmptyEntries))
            {
                int colon = part.IndexOf(':');
                if (colon < 0)
                {
                    continue;
                }

                string key = part[..colon].Trim();
                string val = part[(colon + 1)..].Trim();
                // Skip header segments — a bare pin name has no spaces or brackets
                if (key.Contains(' ') || key.Contains('['))
                {
                    continue;
                }

                if (double.TryParse(val, out double v))
                {
                    dict[key] = v;
                }
            }
            return dict;
        }
    }

    public static class BvLineUtil
    {
        public static Dictionary<string, double> GenDictionary(Func<string> selector)
        {
            var powerNameDic = new Dictionary<string, double>();
            string targetLine = selector();
            string[] spt1 = targetLine.Split([','], StringSplitOptions.RemoveEmptyEntries);
            for (int sptIdx = 2; sptIdx < spt1.Length; sptIdx++)
            {
                string[] arr = spt1[sptIdx].Split(['='], StringSplitOptions.RemoveEmptyEntries);
                string powerName = "";
                double powerVoltage = 0;
                if (arr.Length == 2)
                {
                    powerName = arr[0];
                    _ = double.TryParse(arr[1], out powerVoltage);
                }
                powerNameDic.TryAdd(powerName, powerVoltage);
            }
            return powerNameDic;
        }
    }
}
