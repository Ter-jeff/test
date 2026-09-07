using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.common.ReaderWriter.Reader.InputDataBase;

namespace Cautogen.AutoCZ.CharPostProcessor.InputReader
{
    public class PatListReader
    {
        public static Dictionary<string, PatternData> Read(string fileName)
        {
            //GeneralFunc.WriteMessage("Reading PatList file " + fileName); 

            var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
            var fileReader = new StreamReader(fs);
            var headerOrder = new Dictionary<string, int>();
            var patternList = new Dictionary<string, PatternData>();
            string line = fileReader.ReadLine();
            if (line == null)
            {
                return patternList;
            }

            try
            {
                int index = 0;
                foreach (string str in line.Split(','))
                {
                    headerOrder.Add(str.Replace("\"", "").ToLower(), index);
                    index++;
                }
                while ((line = fileReader.ReadLine()) != null)
                {
                    line = line.Replace("\"", "");
                    if (!Regex.IsMatch(line, @"^[\*|\d].*"))
                    {
                        continue;
                    }

                    var patternData = new PatternData();
                    List<string> lineData = line.Split(',').ToList();
                    int lineCount = lineData.Count;
                    if (lineCount <= index)
                    {
                        for (int i = 0; i < index - lineCount; i++)
                        {
                            lineData.Add("");
                        }
                    }
                    string[] strArray = lineData.ToArray();
                    if (headerOrder.TryGetValue("pattern", out int value))
                    {
                        patternData.PatternName = strArray[value].ToLower();
                    }

                    if (headerOrder.TryGetValue("latest version", out int value1))
                    {
                        patternData.LatestVersion = strArray[value1].ToLower();
                    }

                    if (headerOrder.TryGetValue("use/no use", out int value2))
                    {
                        patternData.Use = strArray[value2].ToLower();
                    }

                    if (headerOrder.TryGetValue("org", out int value3))
                    {
                        patternData.Org = strArray[value3].ToLower();
                    }

                    if (headerOrder.TryGetValue("type spec", out int value4))
                    {
                        patternData.TypeSpec = strArray[value4].ToLower();
                    }

                    if (headerOrder.TryGetValue("timeset version", out int value5))
                    {
                        patternData.TimesetVersion = Path.GetFileNameWithoutExtension(strArray[value5]).ToLower();
                    }
                    else if (headerOrder.TryGetValue("timeset latest", out int value6))
                    {
                        patternData.TimesetVersion = Path.GetFileNameWithoutExtension(strArray[value6]).ToLower();
                    }

                    if (headerOrder.TryGetValue("file versions", out int value7))
                    {
                        patternData.FileVersion = strArray[value7].ToLower();
                    }

                    if (headerOrder.TryGetValue("opcode", out int value8))
                    {
                        patternData.OpCode = strArray[value8].ToLower();
                    }

                    if (headerOrder.TryGetValue("scanmode", out int value9))
                    {
                        patternData.ScanMode = strArray[value9].ToLower();
                    }

                    if (headerOrder.TryGetValue("halt", out int value10))
                    {
                        patternData.Halt = strArray[value10].ToLower();
                    }

                    if (headerOrder.TryGetValue("original timing mode", out int value11))
                    {
                        patternData.OriginalTimingMode = strArray[value11].ToLower();
                    }

                    if (headerOrder.TryGetValue("check", out int value12))
                    {
                        patternData.Check = strArray[value12].ToLower();
                    }

                    if (headerOrder.TryGetValue("t/p category", out int value13))
                    {
                        patternData.TpCategory = strArray[value13].ToLower();
                    }

                    if (!patternList.ContainsKey(patternData.PatternName))
                    {
                        patternList.Add(patternData.PatternName, patternData);
                    }
                }
                return patternList;
            }
            catch (Exception ex)
            {
                throw new Exception("Reading pattern list failed, please check pattern list format. " + ex.Message);
            }
            finally
            {
                fileReader.Close();
                fs.Close();
            }
        }
    }
}
