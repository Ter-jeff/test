using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using CommonReaderLib;

namespace TestPlanLib.Harvest
{
    public partial class DigitalFlagsSheet : MySheet
    {
        [GeneratedRegex(@"\[\d+:\d+\]")]
        private static partial Regex MyRegex();

        public List<DigitalFlagsSheetRow> Rows { set; get; }

        public int IndexIp = -1;
        public int IndexFlagName = -1;
        public int IndexStatement = -1;
        public int IndexPrint = -1;
        public int IndexScanFlag = -1;
        public int IndexMbistFlag = -1;
        public int IndexComments = -1;

        public DigitalFlagsSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }

        public void ExpandCoreRow()
        {
            for (int i = 0; i < Rows.Count; i++)
            {
                DigitalFlagsSheetRow row = Rows[i];
                int expandCount = 0;
                Match match = MyRegex().Match(row.FlagName);
                if (match.Success)
                {
                    bool tryStart = int.TryParse(match.Value.Split(['[', ']', ':'], StringSplitOptions.RemoveEmptyEntries).First(), out int startTmp);
                    bool tryEnd = int.TryParse(match.Value.Split(['[', ']', ':'], StringSplitOptions.RemoveEmptyEntries).Last(), out int endTmp);
                    if (!tryStart || !tryEnd)
                    {
                        continue;
                    }

                    int start = startTmp > endTmp ? startTmp : endTmp;
                    int end = startTmp > endTmp ? endTmp : startTmp;
                    for (int j = start; j >= end; j--)
                    {
                        if (j == end)
                        {
                            row.Ip = row.Ip.Replace(match.Value, j.ToString());
                            row.FlagName = row.FlagName.Replace(match.Value, j.ToString());
                            row.Statement = row.Statement.Replace(match.Value, j.ToString());
                            row.Print = row.Print.Replace(match.Value, j.ToString());
                            row.ScanFlag = row.ScanFlag.Replace(match.Value, j.ToString());
                            row.MbistFlag = row.MbistFlag.Replace(match.Value, j.ToString());
                            row.Comments = row.Comments.Replace(match.Value, j.ToString());
                        }
                        else
                        {
                            var expandRow = new DigitalFlagsSheetRow
                            {
                                Ip = row.Ip.Replace(match.Value, j.ToString()),
                                FlagName = row.FlagName.Replace(match.Value, j.ToString()),
                                Statement = row.Statement.Replace(match.Value, j.ToString()),
                                Print = row.Print.Replace(match.Value, j.ToString()),
                                ScanFlag = row.ScanFlag.Replace(match.Value, j.ToString()),
                                MbistFlag = row.MbistFlag.Replace(match.Value, j.ToString()),
                                Comments = row.Comments.Replace(match.Value, j.ToString())
                            };
                            Rows.Insert(i + expandCount, expandRow);
                            expandCount++;
                        }
                    }
                }
                i += expandCount;
            }
        }

        public void ExportToTxt(string filePath)
        {
            var lines = new List<string>();
            var headerRow = new DigitalFlagsSheetRow
            {
                Ip = "IP",
                FlagName = "Flag Name",
                Statement = "Statement",
                Print = "Print",
                ScanFlag = "SCAN Flag",
                MbistFlag = "MBIST Flag",
                Comments = "Comments"
            };

            Rows.Insert(0, headerRow);

            foreach (DigitalFlagsSheetRow row in Rows)
            {
                string[] rowInfo = [row.Ip, row.FlagName, row.Statement, row.Print, row.ScanFlag, row.MbistFlag, row.Comments];
                string line = string.Join("\t", rowInfo);
                lines.Add(line);
            }

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            File.WriteAllLines(filePath, [.. lines]);

        }
    }
    public class DigitalFlagsSheetRow : MyRow
    {
        public string Ip = "";
        public string FlagName = "";
        public string Statement = "";
        public string Print = "";
        public string ScanFlag = "";
        public string MbistFlag = "";
        public string Comments = "";
        #region Constructor
        public DigitalFlagsSheetRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion

    }
}
