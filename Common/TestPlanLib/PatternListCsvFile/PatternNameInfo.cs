using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TestPlanLib.PatternListCsvFile
{
    public partial class PatternNameInfo
    {
        [GeneratedRegex(@"\..+\\", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"\..+|\.", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@".+/|.+\\", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegex2();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();

        //Only Support N20 Product's Naming Rule!!
        public string OriginalUpperName = "N/A";
        public string FullNameWoMod = "N/A";
        //¥h±¼PatVersion, SiliconVersion, TimeStamp
        public string GenericName = "N/A";
        //PP DD CZ FA
        public string Header = "N/A";
        public string TimeStamp = "N/A";
        //Organization : 'A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP'
        public string Organization = "N/A";
        //TypeSpec : 'AN:HARD_IP,BI:BIST,CH:SCAN,EF:HARD_IP,FU:HARD_IP,IO:HARD_IP,JT:HARD_IP,SC:SCAN,PW:HARD_IP'
        public string TypeSpec = "N/A";
        public string ProjectCode = "N/A";
        //A0 B0
        public string SiliconVersion = "N/A";
        public int PatternVersion;
        public string TpCategory = "N/A";
        public bool IsMod = false;

        public string OpCode = "N/A";

        public string PatternVersionStr
        {
            get { return $"{PatternVersion}_{SiliconVersion}_{TimeStamp}"; }
        }

        public PatternNameInfo(string patName, string regDate = @"\d{10}")
        {
            if (patName == "N/A" || patName.Length == 0)
            {
                return;
            }

            patName = _regex.Replace(patName, "").ToUpper().Trim();
            patName = _regex2.Replace(patName, "").ToUpper().Trim();
            patName = _regex3.Replace(patName, "").ToUpper().Trim();

            OriginalUpperName = patName;
            if (patName.ContainsIgnoreCase("_DM_"))
            {
                OpCode = "DUAL";
            }
            if (patName.ContainsIgnoreCase("_SI_"))
            {
                OpCode = "SINGLE";
            }

            FullNameWoMod = patName;
            Header = patName.Split('_')[0];
            Organization = patName.Split('_').Length > 2 ? patName.Split('_')[2] : "";
            TypeSpec = patName.Split('_').Length > 4 ? patName.Split('_')[4] : "N/A";
            ProjectCode = patName.Split('_').Length > 1 ? patName.Split('_')[1].ToUpper() : "";
            //H:SOC¬OÅv©y¤§­p

            //Organization : 'A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP'
            //TypeSpec : 'AN:HARD_IP,BI:BIST,CH:SCAN,EF:HARD_IP,FU:HARD_IP,IO:HARD_IP,JT:HARD_IP,SC:SCAN,PW:HARD_IP'
            string[] orgAry = "A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP,H:SOC".Split(',');
            foreach (string s in orgAry)
            {
                string org = s.Split(':')[0];
                string cat = s.Split(':')[1];
                if (org == Organization)
                {
                    TpCategory = cat;
                }
            }

            string[] tSpec =
                "AN:HARD_IP,BI:BIST,CH:SCAN,EF:HARD_IP,FU:HARD_IP,IO:HARD_IP,JT:HARD_IP,SC:SCAN,PW:HARD_IP,XX:OTHERS"
                    .Split(',');
            foreach (string s in tSpec)
            {
                string type = s.Split(':')[0];
                string cat = s.Split(':')[1];
                if (type == TypeSpec)
                {
                    if (cat == "HARD_IP")
                    {
                        TpCategory = "HARD_IP";
                    }
                    else if (cat == "OTHERS")
                    {
                        TpCategory = "OTHERS";
                    }
                    else if (TpCategory != "HARD_IP")
                    {
                        //SOC_BIST
                        TpCategory = TpCategory + "_" + cat;
                    }
                }
            }

            string regPattern = $"(?<name>.*_{regDate})";
            if (Regex.IsMatch(patName, regPattern))
            {
                patName = Regex.Match(patName, regPattern).Groups["name"].ToString();
                int len = patName.Split('_').Length;
                if (int.TryParse(patName.Split('_')[len - 3], out int resA))
                {
                    PatternVersion = resA;
                }

                SiliconVersion = patName.Split('_')[len - 2];
                TimeStamp = patName.Split('_')[len - 1];

                if (Regex.IsMatch(TimeStamp, regDate) || TimeStamp.ContainsIgnoreCase("YYMMDDHHMM"))
                {
                    GenericName = Regex.Replace(patName, $"_{PatternVersion}_{SiliconVersion}_{TimeStamp}", "");
                }
            }
            else
            {
                //For DFT Pattern List xxxx_<>
                GenericName = patName;
            }
        }
    }
}
