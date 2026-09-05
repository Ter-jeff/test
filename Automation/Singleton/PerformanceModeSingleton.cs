using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using TestPlanLib.Static;

namespace Automation.Singleton
{
    public class PerformanceModeSingleton
    {
        private static PerformanceModeSingleton _instance;
        private MultiTestSettingSheetsSingleton _testSettings;
        private List<string> _performanceModes;

        private PerformanceModeSingleton()
        {
            if (EpWorkbook.TestPlanWorkbook != null && EpWorkbook.TestPlanWorkbook.Worksheets[NeededSheets.PowerMerge] != null)
            {
                _testSettings = MultiTestSettingSheetsSingleton.Instance();
            }
        }

        public static PerformanceModeSingleton Instance()
        {
            return _instance ?? (_instance = new PerformanceModeSingleton());
        }

        public static string RegContainPerformanceModeWithGroup
        {
            get
            {
                return "(?!Mbist)(?<pmode>M([a-zA-Z]){1}([a-zA-Z0-9]){1}(?<modenumber>[a-fA-F0-9|x|X]{3}))";//MEX001 or MX0001
            }
        }
        public static string RegContainPerformanceModeByPattern
        {
            get
            {
                return "(?!Mbist)^(?<pmode>M([a-zA-Z]){1}([a-zA-Z0-9]){1}(?<modenumber>[a-fA-F0-9|x|X]{2,3}))";//MEX001 or MX0001 or MGXXXX
            }
        }
        public static string RegContainPerformanceMode
        {
            get
            {
                return "(?!Mbist)M([a-zA-Z]){1}([a-zA-Z0-9]){2}([a-zA-Z0-9]){1,2}";     // Update by JN 20200114  to support MGXX01  // 200709 add MGXXXX
            }
        }
        public static string RegPerformanceMode
        {
            get
            {
                return "^" + RegContainPerformanceMode + "$";
            }
        }

        public static void Initialize()
        {
            _instance = null;
        }

        public bool IsPerformanceMode(string input)
        {
            return Regex.IsMatch(input, RegPerformanceMode, RegexOptions.IgnoreCase);
        }

        public string FindPerformanceMode(string input)
        {
            foreach (string str in input.Split('_'))
            {
                if (Regex.IsMatch(str, RegContainPerformanceMode, RegexOptions.IgnoreCase))
                {
                    return Regex.Match(str, RegContainPerformanceMode, RegexOptions.IgnoreCase).Groups[0].ToString();
                }
            }
            return "";
        }

        //Only for new Test PerformanceMode version
        public List<string> GetAllPerformanceMode()
        {
            if (_performanceModes != null)
            {
                return _performanceModes;
            }

            if (_testSettings != null)
            {
                _performanceModes = _testSettings.PerformanceModeList;
                return _performanceModes;
            }

            return new List<string>();
        }

        public Dictionary<string, string> GetAllPerformanceModeDic()
        {
            List<string> performanceModeList = GetAllPerformanceMode();
            var performanceModeDic = new Dictionary<string, string>();
            int dig456 = 3;
            foreach (IGrouping<string, string> group in performanceModeList.GroupBy(x => x.Substring(0, 3)))
            {
                if (group.Count() >= 10)
                {
                    int mod = (group.Count() - 6) % dig456;
                    int step = (group.Count() - 6) / dig456;
                    List<int> maxCnt = new List<int>();
                    for (int i = 0; i < dig456; i++)
                    {
                        maxCnt.Add(i < mod ? step + 1 : step);
                    }

                    int final = 0;
                    for (int index = 0; index < group.ToList().Count; index++)
                    {
                        int cnt = index + 1;
                        string mode = group.ToList()[index].ToUpper();
                        if (Regex.IsMatch(mode, RegContainPerformanceModeWithGroup, RegexOptions.IgnoreCase))
                        {
                            if (cnt <= dig456)
                            {
                                if (!performanceModeDic.ContainsKey(mode))
                                {
                                    performanceModeDic.Add(mode, cnt.ToString());
                                }
                            }
                            else if (cnt > dig456 && cnt <= group.Count() - dig456)
                            {
                                int num = 0;
                                int total = 0;
                                for (int i = 0; i < maxCnt.Count; i++)
                                {
                                    total += maxCnt[i];
                                    if (total >= cnt - dig456)
                                    {
                                        num = i;
                                        break;
                                    }
                                }

                                if (!performanceModeDic.ContainsKey(mode))
                                {
                                    performanceModeDic.Add(mode, (num + 4).ToString());
                                }
                            }
                            else if (cnt > group.Count() - dig456)
                            {

                                if (!performanceModeDic.ContainsKey(mode))
                                {
                                    performanceModeDic.Add(mode, (final + 7).ToString());
                                    final++;
                                }
                            }
                        }
                    }
                }
                else
                {
                    bool flag = false;
                    foreach (string mode in group)
                    {
                        if (Regex.IsMatch(mode, RegContainPerformanceModeWithGroup, RegexOptions.IgnoreCase))
                        {
                            string number = Regex.Match(mode, RegContainPerformanceModeWithGroup, RegexOptions.IgnoreCase).Groups["modenumber"].ToString();
                            if (!int.TryParse(number, out int _))
                            {
                                flag = true;
                                break;
                            }
                        }
                    }

                    if (flag)
                    {
                        int cnt = 1;
                        foreach (string mode in group)
                        {
                            if (!performanceModeDic.ContainsKey(mode.ToUpper()))
                            {
                                performanceModeDic.Add(mode.ToUpper(), cnt.ToString());
                                cnt++;
                            }
                        }
                    }
                    else
                    {
                        foreach (string mode in group)
                        {
                            if (Regex.IsMatch(mode, RegContainPerformanceModeWithGroup, RegexOptions.IgnoreCase))
                            {
                                string number = Regex.Match(mode, RegContainPerformanceModeWithGroup, RegexOptions.IgnoreCase).Groups["modenumber"].ToString();
                                if (!performanceModeDic.ContainsKey(mode.ToUpper()))
                                {
                                    performanceModeDic.Add(mode.ToUpper(), number.Last().ToString());
                                }
                            }
                        }
                    }
                }
            }
            return performanceModeDic;
        }
    }
}
