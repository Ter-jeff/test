using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonReaderLib;

namespace TestPlanLib.BinCut.NonIgxlSheet.HarvestDssc
{
    public static class HarvPinGroupTableReader
    {
        public class HarvPinGroup : MySheet
        {
            public class HarvPinGrpRow : MyRow
            {
                public string Line = "";
                public List<string> RowData = [];
            }

            public string Titlerow = "";
            public int PatternIdx = -1;
            public int PinGroupIdx = -1;
            public int PinFlagIdx = -1;
            public List<HarvPinGrpRow> Rows = [];

            public static HarvPinGroup Read(string inPath)
            {
                List<string> lines = [.. File.ReadAllLines(inPath)];
                HarvPinGroup errors = Read(lines, inPath);
                return errors;
            }

            private static HarvPinGroup Read(List<string> lines, string sheetName)
            {
                try
                {
                    HarvPinGroup harvPinGroupTable = ReadSheet(lines);

                    return harvPinGroupTable;
                }
                catch (Exception e)
                {
                    string msg = "Find exception when reading " + sheetName + "!!! \n ErrMsg: " + e;
                    var errMsg = new Exception(msg);
                    throw errMsg;
                }
            }

            private static HarvPinGroup ReadSheet(List<string> lines)
            {
                var harvPinGroupTable = new HarvPinGroup();
                harvPinGroupTable.Rows.Clear();
                int index = 0;

                for (; index < lines.Count; index++)
                {
                    string line = lines[index];

                    if (line.Contains("PATTERN", StringComparison.OrdinalIgnoreCase))
                    {
                        harvPinGroupTable.Titlerow = line;
                        string[] spt = line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < spt.Length; i++)
                        {
                            if (spt[i].Contains("PATTERN", StringComparison.OrdinalIgnoreCase))
                            {
                                harvPinGroupTable.PatternIdx = i;
                            }
                            else if (spt[i].Contains("HARVPINGRP", StringComparison.OrdinalIgnoreCase))
                            {
                                harvPinGroupTable.PinGroupIdx = i;
                            }
                            else if (spt[i].Contains("HARVPINFLAG", StringComparison.OrdinalIgnoreCase))
                            {
                                harvPinGroupTable.PinFlagIdx = i;
                            }
                        }
                        break;
                    }
                }

                index++;

                for (; index < lines.Count; index++)
                {
                    string line = lines[index];
                    string[] spt = line.Split(['\t'], StringSplitOptions.RemoveEmptyEntries);
                    var vLine = spt.Select(token => token.Trim()).ToList();
                    var harvRow = new HarvPinGrpRow
                    {
                        Line = line,
                        RowData = vLine
                    };
                    harvPinGroupTable.Rows.Add(harvRow);
                }
                return harvPinGroupTable;
            }
        }
    }
}
