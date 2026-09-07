using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;

namespace Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    public class PatternClass
    {
        private const string Pattern = "([,&;])";
        public const string MultipleConst = "Multiple_";
        public const string MtdConst = "MTD_";
        public string TestPlanPatternName = "";
        public string RealPatternName = "";
        public List<string> InstancePatternName = new List<string>();
        public List<string> InstancePayloadName = new List<string>();
        public List<List<string>> PatternSetList = new List<List<string>>();

        public bool IsDefaultNobinout = false;

        public PatternClass(string patternName)
        {
            TestPlanPatternName = patternName.ToLower();
            RealPatternName = patternName;

            foreach (string patternSet in TestPlanPatternName.Split(',', '#'))
            {
                foreach (string pattern in patternSet.Split('+'))
                {
                    InstancePatternName.Add(pattern);
                }
                PatternSetList.Add(patternSet.Split('+').ToList());
            }

            IsDefaultNobinout = HardIpConstData.IgrPatBinOutPrefix.Any(p => GetLastPayload().StartsWith(p, StringComparison.OrdinalIgnoreCase));

        }

        public PatternClass(PatternClass other)
        {
            if (other == null)
            {
                return;
            }

            TestPlanPatternName = other.TestPlanPatternName;
            RealPatternName = other.RealPatternName;

            InstancePatternName = new List<string>(other.InstancePatternName);
            InstancePayloadName = new List<string>(other.InstancePayloadName);

            PatternSetList = other.PatternSetList?.Select(innerList => new List<string>(innerList)).ToList() ?? new List<List<string>>();
        }

        public PatternClass Copy()
        {
            return new PatternClass(this);
        }

        public void ConvertRealPatternName(ScghData scghData)
        {
            if (scghData == null || scghData.ConvertedPatternRowListByHardip.Count == 0)
            {
                return;
            }

            // Init1,Init2,init3 ; payload => A1,A2,A3 & B1,B2,B3 & C1,C2,C3 ; payload
            RealPatternName = "";
            List<string> realPatternNameText = new List<string>();
            bool lastFlag = false;

            var newPatternSetList = new List<List<string>>();

            foreach (List<string> patternSet in PatternSetList)
            {
                foreach (string pattern in patternSet)
                {
                    var patlist = new List<List<string>>();
                    scghData.GetPatternListFromScgh(pattern, ref patlist);
                }
            }

            PatternSetList = newPatternSetList;

            foreach (string pat in Regex.Split(TestPlanPatternName, Pattern).ToList())
            {
                List<string> realPatternNameList = new List<string>();
                List<List<string>> patlist = new List<List<string>>();
                bool flag = scghData.GetPatternListFromScgh(pat, ref patlist);
                foreach (List<string> patternList in patlist)
                {
                    realPatternNameList.Add(string.Join(",", patternList));
                }

                //divide by convertedPatternRowList
                realPatternNameText.Add(string.Join("|", realPatternNameList));

                if (lastFlag && realPatternNameText.Last() == ",")
                {
                    realPatternNameText[realPatternNameText.Count - 1] = "&"; //when using alias , change "," to "&" for creating patset
                }

                lastFlag = flag;
            }
            RealPatternName = string.Join("", realPatternNameText);
        }

        public string GetPatternName()
        {
            if (IsMultiple())
            {
                return MultipleConst + GetLastPayload();
            }

            if (IsMultiTimeDomain())
            {
                return MtdConst + GetMtdLastPayLoad();
            }

            if (HardIpConstData.RegInsInPatt.IsMatch(RealPatternName) ||
                HardIpConstData.RegOpcodeInPatt.IsMatch(RealPatternName))
            {
                return RealPatternName;
            }

            return GetLastPayload();
        }

        public bool IsMultiple()
        {
            return PatternSetList.SelectMany(p => p).Count() > 1;
        }

        public bool IsMultiTimeDomain()
        {
            return RealPatternName.Contains("#");
        }

        public bool IsFullMtdPattern()
        {
            return RealPatternName.Split('+', '#').Length == InstancePatternName.Count;
        }

        public List<string> GetAliasPatternList()
        {
            return Regex.Split(TestPlanPatternName, Pattern).ToList();
        }

        public string GetLastPayload()
        {
            if (IsMultiTimeDomain())
            {
                return GetMtdLastPayLoad();
            }

            if (PatternSetList.Count == 0)
            {
                return "";
            }

            if (PatternSetList.Last().Count == 0)
            {
                return "";
            }

            return PatternSetList.Last().Last();
        }

        public string GetMtdLastPayLoad()
        {
            return InstancePatternName.Last();
        }

        public List<string> GetMtdEachLastPayLoad()
        {
            return RealPatternName.Split('#').Select(pats => pats.Split('+').Last()).ToList();
        }

        public string GetInstancePatternName()
        {
            if (!Regex.IsMatch(TestPlanPatternName, HardIpConstData.NoPattern, RegexOptions.IgnoreCase) &&
                !HardIpConstData.RegInsInPatt.IsMatch(TestPlanPatternName))
            {
                return string.Join(";", InstancePatternName);
            }
            return "";
        }
    }
}
