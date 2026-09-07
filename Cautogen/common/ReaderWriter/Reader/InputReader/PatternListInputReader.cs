using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Cautogen.common.ReaderWriter.Reader.InputDataBase;

using OfficeOpenXml;

namespace Cautogen.common.ReaderWriter.Reader.InputReader
{
    public class PatternListInputReader : TextInputReader
    {
        // data member
        public static Dictionary<string, PatternData> PatternList = new Dictionary<string, PatternData>(StringComparer.CurrentCultureIgnoreCase);

        // constructor
        public PatternListInputReader(string filePath)
            : base(filePath)
        {
            // reset previous result
            PatternList = new Dictionary<string, PatternData>();
        }

        public bool Read(ExcelWorksheet excelWorksheet)
        {
            var headerOrder = new Dictionary<string, int>();

            int startColNumber = excelWorksheet.Dimension.Start.Column;
            int startRowNumber = excelWorksheet.Dimension.Start.Row;
            int endColNumber = excelWorksheet.Dimension.End.Column;
            int endRowNumber = excelWorksheet.Dimension.End.Row;

            try
            {
                // read header
                int index = 0;
                for (int j = startColNumber; j <= endColNumber; j++)
                {
                    if (string.IsNullOrEmpty(excelWorksheet.Cells[startRowNumber, j].Text))
                    {
                        continue;
                    }

                    string cell = excelWorksheet.Cells[startRowNumber, j].Value.ToString();
                    headerOrder.Add(cell.Trim().Replace("\"", "").ToUpper(), index);
                    index++;
                }
                if (!headerOrder.Keys.Contains("PATTERN"))
                {
                    return false;
                }

                // read content
                for (int i = startRowNumber + 1; i <= endRowNumber; i++)
                {
                    var lineData = new List<string>();
                    for (int j = startColNumber; j <= endColNumber; j++)
                    {
                        if (string.IsNullOrEmpty(excelWorksheet.Cells[i, j].Text))
                        {
                            continue;
                        }

                        string cell = excelWorksheet.Cells[i, j].Value.ToString();
                        lineData.Add(cell);
                    }

                    bool isBlankRow = lineData.All(string.IsNullOrEmpty);
                    if (isBlankRow)
                    {
                        continue;
                    }

                    string[] strArray = lineData.ToArray();

                    PatternData patternData = ReadRow(headerOrder, strArray);

                    if (!PatternList.ContainsKey(patternData.PatternName))
                    {
                        PatternList.Add(patternData.PatternName, patternData);
                    }
                }

            }
            catch (Exception e)
            {
                throw new Exception("Reading pattern list failed, may be caused by wrong format of pattern list. " + e.Message);
            }
            if (PatternList.Count > 0)
            {
                return true;
            }

            return false;
        }

        // methods
        protected override void _Read(StreamReader textReader)
        {

            var headerOrder = new Dictionary<string, int>();
            string line = textReader.ReadLine();

            try
            {
                int index = 0;
                if (line == null)
                {
                    return;
                }

                // read header
                foreach (string str in line.Split(','))
                {
                    if (string.IsNullOrEmpty(str))
                    {
                        continue;
                    }

                    headerOrder.Add(str.Trim().Replace("\"", "").ToUpper(), index);
                    index++;
                }

                // read content
                while ((line = textReader.ReadLine()) != null)
                {
                    line = line.Replace("\"", "");
                    if (line.Trim() == "")
                    {
                        continue;
                    }

                    List<string> lineData = line.Split(',').ToList();
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
                        if (lineData[i].Trim() == "")
                        {
                            continue;
                        }

                        isBlankRow = false;
                        break;
                    }
                    if (isBlankRow)
                    {
                        continue;
                    }

                    string[] strArray = lineData.ToArray();

                    PatternData patternData = ReadRow(headerOrder, strArray);

                    if (!PatternList.ContainsKey(patternData.PatternName))
                    {
                        PatternList.Add(patternData.PatternName, patternData);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("Reading pattern list failed, may be caused by wrong format of pattern list. " + e.Message);
            }
        }

        private static PatternData ReadRow(IDictionary<string, int> headerOrder, IList<string> strArray)
        {
            var patternData = new PatternData();

            if (headerOrder.TryGetValue("PATTERN", out int value))
            {
                patternData.PatternName = strArray[value].ToUpper();
            }

            if (headerOrder.TryGetValue("PATTERN NAME", out int value1))
            {
                patternData.PatternName = strArray[value1].ToUpper();
            }

            if (patternData.PatternName.Trim() == "")
            {
                return patternData;
            }

            if (headerOrder.TryGetValue("LATEST VERSION", out int value2))
            {
                patternData.LatestVersion = strArray[value2].ToLower();
            }

            if (headerOrder.TryGetValue("USE/NO USE", out int value3))
            {
                patternData.Use = strArray[value3].ToLower();
            }

            if (headerOrder.TryGetValue("ORG", out int value4))
            {
                patternData.Org = strArray[value4].ToLower();
            }

            if (headerOrder.TryGetValue("TYPE SPEC", out int value5))
            {
                patternData.TypeSpec = strArray[value5].ToLower();
            }

            if (headerOrder.ContainsKey("TIMESET LATEST"))
            {
                if (strArray[headerOrder["TIMESET LATEST"]].Equals("N/A", StringComparison.OrdinalIgnoreCase))
                {
                    patternData.TimesetVersion = "N/A";
                }
                else
                {
                    patternData.TimesetVersion =
                    Path.GetFileNameWithoutExtension(strArray[headerOrder["TIMESET LATEST"]]);
                }
            }

            if (headerOrder.ContainsKey("TIMESET VERSION"))
            {
                if (strArray[headerOrder["TIMESET VERSION"]].Equals("N/A", StringComparison.OrdinalIgnoreCase))
                {
                    patternData.TimesetVersion = "N/A";
                }
                else
                {
                    patternData.TimesetVersion =
                        Path.GetFileNameWithoutExtension(strArray[headerOrder["TIMESET VERSION"]]);
                }
            }

            if (headerOrder.TryGetValue("FILE VERSIONS", out int value6))
            {
                patternData.FileVersion = strArray[value6].ToLower();
            }

            if (headerOrder.TryGetValue("OPCODE", out int value7))
            {
                patternData.OpCode = strArray[value7].ToLower();
            }

            if (headerOrder.TryGetValue("SCANMODE", out int value8))
            {
                patternData.ScanMode = strArray[value8].ToLower();
            }

            if (headerOrder.TryGetValue("HALT", out int value9))
            {
                patternData.Halt = strArray[value9].ToLower();
            }

            if (headerOrder.TryGetValue("ORIGINAL TIMING MODE", out int value10))
            {
                patternData.OriginalTimingMode = strArray[value10].ToLower();
            }

            if (headerOrder.TryGetValue("CHECK", out int value11))
            {
                patternData.Check = strArray[value11].ToLower();
            }

            if (headerOrder.TryGetValue("T/P CATEGORY", out int value12))
            {
                patternData.TpCategory = strArray[value12].ToLower();
            }

            return patternData;
        }
    }
}

