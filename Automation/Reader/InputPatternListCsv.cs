using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

using TestPlanLib.PatternListCsvFile;

namespace Automation.Reader
{
    public class InputPatternListCsv
    {
        public string FullName { get; private set; }

        public InputPatternListCsv(FileInfo fileInfo)
        {
            FullName = fileInfo.FullName;
        }

        public bool Compile(string compilePatFileName, string patternPath, string timeSetPath, bool useLatestVersion, out Dictionary<string, CompileItem> dicCompileItem, Action<string, EnumMessageLevel, int> writeMessage)
        {
            dicCompileItem = null;

            Response.Report("Parsing Pattern Dashboard ...", EnumMessageLevel.General);

            if (!Directory.Exists(timeSetPath))
            {
                Response.Report("Please check timeset path: " + timeSetPath, EnumMessageLevel.Error);

                return false;
            }

            var patternList = PatternListSingleton.GetInstance(FullName, timeSetPath);
            if (patternList.IsCompiledPatternDashboard())
            {
                Response.Report("Pattern dashboard use the format that already includes CompiledPat.csv information.", EnumMessageLevel.General);

                return true;
            }

            if (patternList.HeaderFailFlag)
            {
                Response.Report("Pattern dashboard parsing failed, Please make sure your header exist. \r\nMust have header:#,Pattern,USE/No Use,Timeset Latest,File Versions", EnumMessageLevel.Error);

                Response.Report("Pattern dashboard retrieval was terminated.", EnumMessageLevel.Error);

                return false;
            }

            if (File.Exists(compilePatFileName))
            {
                //Read Compile File
                var compileReader = new CompilePatReader();
                dicCompileItem = compileReader.ReadCompileFile(compilePatFileName);

                // Update Pattern list by latest pattern version
                if (useLatestVersion)
                {
                    Response.Report("Enable use lastet version pattern of pattern server.", EnumMessageLevel.Info);

                    Dictionary<string, string> latestGenericPatDict = compileReader.GetLatestPatDict();
                    List<string> modPatStrings = patternList.UpdateCsvByLatestPattern(latestGenericPatDict);

                    if (modPatStrings != null)
                    {
                        foreach (string modPatStr in modPatStrings)
                        {
                            Response.Report(modPatStr, EnumMessageLevel.Warning, -1);
                        }
                    }
                }

                //Merge PatternList
                patternList.UpdatePatternDashboardWithCompiledPatCsv(dicCompileItem);
                FullName = patternList.CompiledPatternDashboardFile;
                Response.Report("Pattern dashboard is not in compiled format. Updated with CompiledPat.csv.", EnumMessageLevel.Info);

                Response.Report(FullName, EnumMessageLevel.Info);

                return true;
            }

            Response.Report("CompiledInfo.csv not found. Please check the input document.", EnumMessageLevel.Error);

            Response.Report("Pattern dashboard retrieval was terminated.", EnumMessageLevel.General);

            return false;
        }

        public Dictionary<string, Tuple<string, string>> GetPatternTimeSet(string timeSetPath)
        {
            var fullPatNameTimeSetDict = new Dictionary<string, Tuple<string, string>>();
            Dictionary<string, OriPatListItem> patternListDic = PatternListSingleton.GetInstance(FullName, LocalSpecs.TimeSetFolder).GetPatternData();
            foreach (KeyValuePair<string, OriPatListItem> pair in patternListDic)
            {
                string patName = pair.Value.FileVersions.ToUpper();
                if ((patName.IndexOf("/", StringComparison.Ordinal) == -1 && patName.IndexOf(".", StringComparison.Ordinal) == -1) || Regex.IsMatch(patName, "N/A", RegexOptions.IgnoreCase))
                {
                    continue;
                }
                patName = patName.Substring(patName.LastIndexOf("/", StringComparison.Ordinal) + 1, patName.Length - patName.LastIndexOf("/", StringComparison.Ordinal) - 1);
                patName = patName.Substring(0, patName.IndexOf(".", StringComparison.Ordinal));

                // USE/No Use
                string useNoUse = pair.Value.UseNoUse;

                // timeSet
                string timeSet = pair.Value.TimeSetVersion;
                if (timeSet.IndexOf(".", StringComparison.Ordinal) != -1)
                { // remove ".TXT"
                    timeSet = timeSet.Substring(0, timeSet.LastIndexOf(".", StringComparison.Ordinal));
                }
                if (!fullPatNameTimeSetDict.ContainsKey(patName) && patName != "")
                {
                    fullPatNameTimeSetDict.Add(patName, new Tuple<string, string>(useNoUse, timeSet));
                }
            }
            return fullPatNameTimeSetDict;
        }

        public string GetProjectName()
        {
            string fileName = Path.GetFileName(FullName);
            if (fileName != null)
            {
                return fileName.Split('_').First();
            }

            return "";
        }
    }
}
