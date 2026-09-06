using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using TestPlanLib.Basic;

namespace Automation.Singleton
{
    public class AcTSetCategoryMapSingleton
    {
        private static AcTSetCategoryMapSingleton _instance;
        private static string _timeSetFolder;
        private static bool _skipPatCheck;
        private static readonly object _locker = new object();
        public readonly Dictionary<string, PatternData> PatternList = new Dictionary<string, PatternData>(StringComparer.OrdinalIgnoreCase);
        public readonly List<TimeSetBlock2Category> TimeSetBlock2Categories = new List<TimeSetBlock2Category>();
        public readonly Dictionary<string, int> DicTimeSetVersion = new Dictionary<string, int>();
        public readonly List<string> AllTimeSetVersionInK;
        public List<ComTimeSetBasicSheet> MultiTimeSetSheets = new List<ComTimeSetBasicSheet>();

        //For patten check
        public HashSet<string> PatternPathListInK = new HashSet<string>();
        private readonly HashSet<string> _patternNameListInK = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
        private static readonly Regex _regex = new Regex(@"_\d+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex2 = new Regex(@"(?<str>.*)_\d+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex3 = new Regex(@".*_(?<ver>\d+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex5 = new Regex(".gz$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex6 = new Regex(".pat$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _regex7 = new Regex(".patx$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly HashSet<string> _searchPatternExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".gz",
            ".pat",
            ".patx"
        };

        public static AcTSetCategoryMapSingleton Instance(string path = "")
        {
            if (_instance == null)
            {
                lock (_locker)
                {
                    if (_instance == null)
                    {
                        _instance = new AcTSetCategoryMapSingleton(path);
                    }
                }
            }
            return _instance;
        }

        public AcTSetCategoryMapSingleton(string path = "")
        {
            List<string> timeSetList = string.IsNullOrEmpty(_timeSetFolder) || !Directory.Exists(_timeSetFolder) ?
                new List<string>() : Directory.GetFiles(_timeSetFolder, "TIMESET*.txt", SearchOption.TopDirectoryOnly).ToList();
            AllTimeSetVersionInK = timeSetList.Select(Path.GetFileNameWithoutExtension).ToList().Distinct().ToList();
            DicTimeSetVersion.Clear();
            foreach (string file in AllTimeSetVersionInK)
            {
                if (_regex.IsMatch(file))
                {
                    string timeSet = _regex2.Match(file).Groups["str"].ToString().ToUpper();
                    int paraVer = Convert.ToInt32(_regex3.Match(file).Groups["ver"].ToString());

                    if (!DicTimeSetVersion.ContainsKey(timeSet))
                    {
                        DicTimeSetVersion.Add(timeSet, paraVer);
                    }
                    else
                    {
                        if (DicTimeSetVersion[timeSet] < paraVer)
                        {
                            DicTimeSetVersion[timeSet] = paraVer;
                        }
                    }
                }
            }

            if (File.Exists(LocalSpecs.PatternListCsvFileName))//TW
            {
                PatternList = PatternListReader.GetPatternListDic(LocalSpecs.PatternListCsvFileName);
                PatternList = ModifyTimeSet(PatternList, DicTimeSetVersion);
            }
            else if (File.Exists(path))
            {
                PatternList = PatternListReader.GetPatternListDic(path);
                PatternList = ModifyTimeSet(PatternList, DicTimeSetVersion);
            }

            string patternPath = LocalSpecs.GetPatternPath();
            if (!string.IsNullOrEmpty(patternPath) && Directory.Exists(patternPath) && !_skipPatCheck)
            {
                var allFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (string file in Directory.EnumerateFiles(patternPath, "*", SearchOption.AllDirectories))
                {
                    if (_searchPatternExtensions.Contains(Path.GetExtension(file)))
                    {
                        allFiles.Add(file);
                    }
                }

                PatternPathListInK = allFiles;
                foreach (string pattern in PatternPathListInK)
                {
                    string name = _regex5.Replace(pattern, "");
                    name = Path.GetFileName(_regex6.Replace(name, ""));
                    name = Path.GetFileName(_regex7.Replace(name, ""));
                    _patternNameListInK.Add(name.ToLower());
                }
            }
            foreach (KeyValuePair<string, PatternData> patten in PatternList)
            {
                if (patten.Value.Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (patten.Value.FileVersion.Equals("n/a", StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }

                if (_patternNameListInK.Contains(patten.Value.PatternVersion) || _skipPatCheck)
                {
                    patten.Value.IsExist = true;
                }
            }
        }

        private Dictionary<string, PatternData> ModifyTimeSet(Dictionary<string, PatternData> patternList, Dictionary<string, int> dicTimeSetVersion)
        {
            foreach (KeyValuePair<string, PatternData> pattern in patternList)
            {
                string fileName = Path.GetFileNameWithoutExtension(pattern.Value.TimesetLatest);
                if (string.IsNullOrEmpty(pattern.Value.TimesetLatest))
                {
                    if (dicTimeSetVersion.ContainsKey(patternList[pattern.Key].TimeSetVersion))
                    {
                        patternList[pattern.Key].TimeSetVersion = patternList[pattern.Key].TimeSetVersion + "_" +
                                                                  dicTimeSetVersion[patternList[pattern.Key].TimeSetVersion];
                    }
                }
                else
                {
                    if (dicTimeSetVersion.TryGetValue(fileName, out int value))
                    {
                        patternList[pattern.Key].TimeSetVersion = fileName + "_" +
                                                                  value;
                    }
                    else
                    {
                        patternList[pattern.Key].TimeSetVersion = pattern.Value.TimesetLatest;
                    }
                }
            }
            return patternList;
        }

        public static void Initialize(string timeSetFolder, bool skipPatCheck)
        {
            _instance = null;
            _timeSetFolder = timeSetFolder;
            _skipPatCheck = skipPatCheck;
        }

        public string CheckAllPatternExist(string pattern)
        {
            string env = "";
            string key = pattern.ToLower();
            if (!PatternList.TryGetValue(key, out PatternData patternData))
            {
                env = "MissPatternInCsv";
            }
            else
            {
                if (patternData.Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
                {
                    env = "dont_useInCsv";
                }
                else
                {
                    if (patternData.FileVersion.Equals("n/a", StringComparison.OrdinalIgnoreCase))
                    {
                        env = "no_patternInCsv";
                    }
                    else
                    {
                        bool checkPattern = !patternData.IsExist;
                        if (checkPattern)
                        {
                            env = "no_pattern";
                        }
                    }
                }
            }
            return env;
        }

        public void SetMultiTimeSetSheet(List<ComTimeSetBasicSheet> timeSetSheets)
        {
            MultiTimeSetSheets = timeSetSheets;
        }

        public string GetCategory(string timeSetVersion)
        {
            string retCategory = "TBD";
            foreach (TimeSetBlock2Category oneRow in TimeSetBlock2Categories)
            {
                if (oneRow.TimeSetSheet.Equals(timeSetVersion, StringComparison.OrdinalIgnoreCase))
                {
                    retCategory = oneRow.Category;
                    return retCategory;
                }
            }
            return retCategory;
        }

        //search AC category with block, NA do not need to search. If exceed 1 more candidate, return TBD.
        public string GetCategory(string tSetSheet, BlockType block)
        {
            string retCategory = "TBD";
            var list = new List<string>();
            if (string.IsNullOrEmpty(tSetSheet))
            {
                return retCategory;
            }
            foreach (TimeSetBlock2Category oneRow in TimeSetBlock2Categories)
            {
                if (!block.Equals(oneRow.Block))
                {
                    continue;
                }

                if (tSetSheet.Split(',').ToList().Exists(p => p.Equals(oneRow.TimeSetSheet, StringComparison.OrdinalIgnoreCase)))
                {
                    list.Add(oneRow.Category);
                }
            }
            if (list.Count == 1)
            {
                retCategory = list[0];
            }

            return retCategory;
        }

        public bool Contains(string tSetSheet)
        {
            bool isFoundSameRow = false;
            foreach (TimeSetBlock2Category oneRow in TimeSetBlock2Categories)
            {
                if (oneRow.TimeSetSheet == tSetSheet)
                {
                    isFoundSameRow = true;
                    break;
                }
            }
            return isFoundSameRow;
        }

        public bool TryContains(string tSetSheet, out TimeSetBlock2Category value)
        {
            value = TimeSetBlock2Categories.FirstOrDefault(row => row.TimeSetSheet == tSetSheet);
            return value != null;
        }

        public string GetTimeSetVersion(string timeSet)
        {
            if (AllTimeSetVersionInK.Exists(x => x.Equals(timeSet, StringComparison.CurrentCultureIgnoreCase)))
            {
                return AllTimeSetVersionInK.Find(x => x.Equals(timeSet, StringComparison.CurrentCultureIgnoreCase));
            }

            timeSet = _regex.Replace(timeSet, "");
            foreach (KeyValuePair<string, int> oneRow in DicTimeSetVersion)
            {
                if (oneRow.Key.Equals(timeSet, StringComparison.OrdinalIgnoreCase))
                {
                    return oneRow.Key + "_" + oneRow.Value;
                }
            }
            return "";
        }

        public bool SetRow(string tSetSheet, BlockType block, string category)
        {
            bool isFoundSameRow = false;
            foreach (TimeSetBlock2Category oneRow in TimeSetBlock2Categories)
            {
                if (oneRow.TimeSetSheet == tSetSheet && oneRow.Block == block && oneRow.Category == category)
                {
                    isFoundSameRow = true;
                    break;
                }
            }

            if (!isFoundSameRow)
            {
                TimeSetBlock2Category oneRow = new TimeSetBlock2Category
                {
                    TimeSetSheet = tSetSheet,
                    Block = block,
                    Category = category
                };
                TimeSetBlock2Categories.Add(oneRow);
            }
            return !isFoundSameRow;
        }

        public string GetCategoryUsageTSetSheetName(string category)
        {
            var tSetSheetList = new List<string>();
            TimeSetBlock2Category[] categories = TimeSetBlock2Categories.FindAll(t => t.Category.Equals(category)).ToArray();
            foreach (TimeSetBlock2Category lItem in categories)
            {
                if (!tSetSheetList.Contains(lItem.TimeSetSheet))
                {
                    tSetSheetList.Add(lItem.TimeSetSheet);
                }
            }
            if (tSetSheetList.Count > 0)
            {
                return string.Join(", ", tSetSheetList.ToArray());
            }

            return "";
        }
    }

    public class TimeSetBlock2Category
    {
        public string TimeSetSheet;
        public BlockType Block;
        public string Category;
    }

    public enum BlockType
    {
        Common,
        Efuse,
        Scan,
        Mbist,
        HardIp,
        BScan,
        Otp,
        SPI_ROM,
        None
    }
}
