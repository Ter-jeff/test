using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;

using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

namespace Cautogen.AutoCZ.CharPostProcessor.AC
{
    public class AcTSetCategoryMapSingleton
    {
        private static AcTSetCategoryMapSingleton _instance;
        private static readonly object _locker = new object();
        public readonly Dictionary<string, PatternData> PatternList = new Dictionary<string, PatternData>(StringComparer.OrdinalIgnoreCase);
        public readonly List<TimeSetBlock2Category> TimeSetBlock2Categories = new List<TimeSetBlock2Category>();
        public readonly Dictionary<string, int> DicTimeSetVersion = new Dictionary<string, int>();
        public readonly List<string> AllTimeSetVersionInK = new List<string>();
        public List<ComTimeSetBasicSheet> MultiTimesetSheets = new List<ComTimeSetBasicSheet>();

        //For patten check
        public HashSet<string> PatternPathListInK;
        private readonly HashSet<string> _patternNameListInK = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
        private List<string> _failPatternList = new List<string>();

        #region Singleton
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
        #endregion

        #region Constructor
        public AcTSetCategoryMapSingleton(string path = "")
        {
            //if (File.Exists(LocalSpecs.PatternFolder))//TW
            //{
            //    var patternListReader = new PatternListReader();
            //    PatternList = patternListReader.GetPatternListDic(LocalSpecs.PatListFile);
            //}
            //else if (File.Exists(path))
            //{
            //    var patternListReader = new PatternListReader();
            //    PatternList = patternListReader.GetPatternListDic(path);
            //}

            PatternList = LocalSpecs.PatternDatas;

            string patternPath = LocalSpecs.PatternFolder;
            if (!string.IsNullOrEmpty(LocalSpecs.PatternFolder) && Directory.Exists(LocalSpecs.PatternFolder))
            {
                PatternPathListInK = new HashSet<string>(Directory.GetFiles(patternPath, "*.gz", SearchOption.AllDirectories));
                foreach (string pattern in PatternPathListInK)
                {
                    string name = Regex.Replace(pattern, ".gz$", "", RegexOptions.IgnoreCase);
                    name = Path.GetFileName(Regex.Replace(name, ".atp$", "", RegexOptions.IgnoreCase));
                    name = Path.GetFileName(Regex.Replace(name, ".pat$", "", RegexOptions.IgnoreCase));
                    name = Path.GetFileName(Regex.Replace(name, ".patx$", "", RegexOptions.IgnoreCase));
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

                if (_patternNameListInK.Contains(patten.Value.GetPatVersion()))
                {
                    patten.Value.IsExist = true;
                }
            }

            //Get all timeset from K folder
            if (!Directory.Exists(LocalSpecs.TimeSetFolder))
            {
                return;
            }

            string[] timeSetList = Directory.GetFiles(LocalSpecs.TimeSetFolder, "TIMESET*.txt", SearchOption.TopDirectoryOnly);
            AllTimeSetVersionInK = timeSetList.Select(Path.GetFileNameWithoutExtension).ToList().Distinct().ToList();
            DicTimeSetVersion.Clear();
            foreach (string file in AllTimeSetVersionInK)
            {
                if (Regex.IsMatch(file, @"_\d+$", RegexOptions.IgnoreCase))
                {
                    string timeset = Regex.Match(file, @"(?<str>.*)_\d+$", RegexOptions.IgnoreCase).Groups["str"].ToString().ToUpper();
                    int paraVer = Convert.ToInt32(Regex.Match(file, @".*_(?<ver>\d+)$", RegexOptions.IgnoreCase).Groups["ver"].ToString());

                    if (!DicTimeSetVersion.ContainsKey(timeset))
                    {
                        DicTimeSetVersion.Add(timeset, paraVer);
                    }
                    else
                    {
                        if (DicTimeSetVersion[timeset] < paraVer)
                        {
                            DicTimeSetVersion[timeset] = paraVer;
                        }
                    }
                }
            }
        }

        #endregion
        public static void Initialize()
        {
            _instance = null;
        }

        //public string CheckAllPatternExist(string pattern)
        //{
        //    var env = "";
        //    if (RetrievePlans.SkipCsvDoc)
        //        return "IgnorePatternListCsv";

        //    var key = pattern.ToLower();
        //    if (!PatternList.ContainsKey(key))
        //        env = "MissPatternInCsv";
        //    else
        //    {
        //        var patternData = PatternList[key];
        //        if (patternData.Use.Equals("dont_use", StringComparison.OrdinalIgnoreCase))
        //        {
        //            env = "dont_useInCsv";
        //        }
        //        else
        //        {
        //            if (patternData.FileVersion.Equals("n/a", StringComparison.OrdinalIgnoreCase))
        //            {
        //                env = "no_patternInCsv";
        //            }
        //            else
        //            {
        //                var checkPattern = !patternData.IsExist;
        //                var checkTimeOfPattern = !Instance().CheckTimeOfPattern(pattern);
        //                if (checkPattern && checkTimeOfPattern)
        //                {
        //                    env = "no_pattern,TimeSet_iusse";
        //                }
        //                else if (checkPattern)
        //                {
        //                    env = "no_pattern";
        //                }
        //                else if (checkTimeOfPattern)
        //                {
        //                    env = "TimeSet_iusse";
        //                }
        //            }
        //        }
        //    }
        //    return env;
        //}

        public void SetfailPatternList(List<string> failPatternList)
        {
            _failPatternList = failPatternList;
        }

        public void SetMultiTimeSetSheet(List<ComTimeSetBasicSheet> timeSetSheets)
        {
            MultiTimesetSheets = timeSetSheets;
        }

        public bool CheckTimeOfPattern(string filename)
        {
            if (_failPatternList.Exists(x => x.Equals(filename, StringComparison.CurrentCultureIgnoreCase)))
            {
                return false;
            }

            return true;
        }

        public string GetTimeSet(string patternName)
        {
            if (PatternList.ContainsKey(patternName.ToLower()))
            {
                return PatternList[patternName.ToLower()].TimesetVersion;
            }
            return "TBD";
        }

        public string GetTimeSet(List<string> patterns)
        {
            var timeSets = new List<string>();
            for (int index = 0; index < patterns.Count; index++)
            {
                string pattern = patterns[index];
                if (string.IsNullOrEmpty(pattern))
                {
                    continue;
                }

                if (PatternList.ContainsKey(pattern.ToLower()))
                {
                    timeSets.Add(PatternList[pattern.ToLower()].TimesetVersion);
                }
            }
            return string.Join(",", timeSets.Distinct());
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
            var targetCand = new List<string>();
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
                    targetCand.Add(oneRow.Category);
                }
                //if (oneRow.TimeSetSheet.Equals(tSetSheet, StringComparison.OrdinalIgnoreCase) && block.Equals(oneRow.Block))
                //{
                //    target_cand.Add(oneRow.Category);
                //}
            }
            if (targetCand.Count == 1)
            {
                retCategory = targetCand[0];
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
                    isFoundSameRow = true; //dont care about block is Efuse or not here
                    break;
                }
            }
            return isFoundSameRow;
        }

        //No use function
        public string GetTimeSetVersion(string timeset)
        {
            if (AllTimeSetVersionInK.Exists(x => x.Equals(timeset, StringComparison.CurrentCultureIgnoreCase)))
            {
                return AllTimeSetVersionInK.Find(x => x.Equals(timeset, StringComparison.CurrentCultureIgnoreCase));
            }

            timeset = Regex.Replace(timeset, @"_\d+$", "");
            foreach (KeyValuePair<string, int> oneRow in DicTimeSetVersion)
            {
                if (oneRow.Key.Equals(timeset, StringComparison.OrdinalIgnoreCase))
                {
                    return oneRow.Key + "_" + oneRow.Value;
                }
            }
            return "";
        }

        public bool CheckTimeSetIsInKfolder(string timeset)
        {
            timeset = Regex.Replace(timeset, @"_\d+$", "");
            if (DicTimeSetVersion.ContainsKey(timeset))
            {
                return true;
            }

            return false;
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
                TimeSetBlock2Category oneRow;
                oneRow.TimeSetSheet = tSetSheet;
                oneRow.Block = block;
                oneRow.Category = category;
                TimeSetBlock2Categories.Add(oneRow);
            }
            return !isFoundSameRow;
        }

        public string GetCategoryUsageTsetSheetName(string category)
        {
            var tSetSheetList = new List<string>();
            TimeSetBlock2Category[] blockCategorys = TimeSetBlock2Categories.FindAll(t => t.Category.Equals(category)).ToArray();
            foreach (TimeSetBlock2Category lItem in blockCategorys)
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

    public struct TimeSetBlock2Category
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
        None
    }
}
