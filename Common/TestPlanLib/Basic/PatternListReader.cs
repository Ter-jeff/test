using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using CommonLib.Enums;

using LogLib.Static;

namespace TestPlanLib.Basic
{
    public partial class PatternListReader
    {
        [GeneratedRegex(".atp$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(".gz$", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();

        private static Dictionary<string, PatternData> ReadPatList(string fileName)
        {
            var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
            var fileReader = new StreamReader(fs);
            var headerOrder = new Dictionary<string, int>();
            var patternList = new Dictionary<string, PatternData>();
            string? line = fileReader.ReadLine();
            int rowIndex = 0;
            try
            {
                int index = 0;
                foreach (string str in line!.Split(','))
                {
                    headerOrder.Add(str.Replace("\"", ""), index);
                    index++;
                }
                while ((line = fileReader.ReadLine()) != null)
                {
                    rowIndex++;
                    line = line.Replace("\"", "");
                    if (line.Trim()?.Length == 0)
                    {
                        Response.Report("Blank Row " + rowIndex + " In Patsetcsv " + fileName + " is skipped.", EnumMessageLevel.Warning, 100);
                        continue;
                    }

                    List<string> lineData = [.. line.Split(',')];
                    int lineCount = lineData.Count;
                    if (lineCount <= index)
                    {
                        for (int i = 0; i < index - lineCount; i++)
                        {
                            lineData.Add("");
                        }
                    }

                    bool isBlankRow = true;
                    for (int i = 0; i < index; i++)
                    {
                        if (lineData[i].Trim()?.Length != 0)
                        {
                            isBlankRow = false;
                            break;
                        }
                    }

                    if (isBlankRow)
                    {
                        Response.Report("Blank Row " + rowIndex + " In Patsetcsv " + fileName + " is skipped.", EnumMessageLevel.Warning, 100);
                        continue;
                    }

                    string[] strArray = [.. lineData];
                    string patternName = headerOrder.TryGetValue("Pattern", out int value) ? strArray[value].ToLower() : "";
                    if (patternName.Trim()?.Length == 0)
                    {
                        Response.Report("Beacuse Pattern is Blank, Row " + rowIndex + " Content " + line + " In Patsetcsv " + fileName + " is skipped.", EnumMessageLevel.Warning, 100);
                        continue;
                    }

                    PatternData patternData = GetPatternData(headerOrder, strArray);

                    patternList.TryAdd(patternData.PatternName, patternData);
                }
                return patternList;
            }
            catch (Exception e)
            {
                throw new Exception("Reading pattern list failed, may be caused by wrong format of pattern list. " + e.StackTrace);
            }
            finally
            {
                fileReader.Close();
                fs.Close();
            }
        }

        private static PatternData GetPatternData(Dictionary<string, int> headerOrder, string[] strArray)
        {
            var patternData = new PatternData();
            if (headerOrder.TryGetValue("Pattern", out int value))
            {
                patternData.PatternName = strArray[value].ToLower();
            }

            if (headerOrder.TryGetValue("Latest Version", out int value1))
            {
                patternData.LatestVersion = strArray[value1].ToLower();
            }

            if (headerOrder.TryGetValue("USE/No Use", out int value2))
            {
                patternData.Use = strArray[value2].ToLower();
            }

            if (headerOrder.TryGetValue("Org", out int value3))
            {
                patternData.Org = strArray[value3].ToLower();
            }

            if (headerOrder.TryGetValue("Type Spec", out int value4))
            {
                patternData.TypeSpec = strArray[value4].ToLower();
            }

            if (headerOrder.TryGetValue("Timeset Latest", out int value5))
            {
                patternData.TimeSetVersion = Path.GetFileNameWithoutExtension(strArray[value5]);
            }
            if (headerOrder.TryGetValue("Timeset Version", out int value6))
            {
                patternData.TimeSetVersion = Path.GetFileNameWithoutExtension(strArray[value6]);
            }

            if (headerOrder.TryGetValue("File Versions", out int value7))
            {
                patternData.FileVersion = strArray[value7].ToLower();
                patternData.PatternVersion = Path.GetFileName(MyRegex().Replace(MyRegex1().Replace(patternData.FileVersion, ""), ""));
            }

            if (headerOrder.TryGetValue("OpCode", out int value8))
            {
                patternData.OpCode = strArray[value8].ToLower();
            }

            if (headerOrder.TryGetValue("ScanMode", out int value9))
            {
                patternData.ScanMode = strArray[value9].ToLower();
            }

            if (headerOrder.TryGetValue("Halt", out int value10))
            {
                patternData.Halt = strArray[value10].ToLower();
            }

            if (headerOrder.TryGetValue("Original Timing Mode", out int value11))
            {
                patternData.OriginalTimingMode = strArray[value11].ToLower();
            }

            if (headerOrder.TryGetValue("Check", out int value12))
            {
                patternData.Check = strArray[value12].ToLower();
            }

            if (headerOrder.TryGetValue("CheckComment", out int value13))
            {
                patternData.CheckComment = strArray[value13].ToLower();
            }

            if (headerOrder.TryGetValue("T/P Category", out int value14))
            {
                patternData.TpCategory = strArray[value14].ToLower();
            }

            if (headerOrder.TryGetValue("ScanSetupTSet", out int value15))
            {
                patternData.ScanTset = strArray[value15].ToLower();
            }

            return patternData;
        }

        public static Dictionary<string, PatternData>? GetPatternListDic(string fileName)
        {
            if (File.Exists(fileName))
            {
                return ReadPatList(fileName);
            }

            return null;
        }
    }
}
