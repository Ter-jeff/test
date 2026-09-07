using System;
using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using TagDiff.Core.Common;
using TagDiff.Core.Static;

namespace TagDiff.Core.Input
{
    internal class FlowReader
    {
        public static List<List<FlowRowCompare>> ReadTxt(string file, string sheet, string job)
        {
            var groupedData = new List<List<FlowRowCompare>>();
            if (!File.Exists(file))
            {
                return groupedData;
            }

            try
            {
                string fullName = Path.GetFileName(file);
                using (var reader = new StreamReader(file))
                {
                    bool foundHeader = false;
                    int lineNumber = 0;

                    while (!reader.EndOfStream)
                    {
                        string? line = reader.ReadLine();
                        lineNumber++;
                        if (!TagDiffMain.IncludingBackup)
                        {
                            if (string.IsNullOrWhiteSpace(line))
                            {
                                break;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(line))
                            {
                                continue;
                            }
                        }

                        string[] columns = line.Split('\t');
                        if (!foundHeader)
                        {
                            if (line.Contains("Opcode") && line.Contains("Parameter") && line.Contains("TName"))
                            {
                                foundHeader = true;
                            }
                            continue;
                        }

                        if (columns.Length < 10)
                        {
                            break;
                        }

                        FlowRowCompare row = GenFlowRowCompare(columns, fullName, sheet, lineNumber);

                        bool flag1 = IsMatchJob(job, row.Job);
                        // Only add if Opcode is not "nop"
                        if (!row.Opcode.EqualsIgnoreCase("nop") && flag1)
                        {
                            groupedData.Add([row]);
                        }
                    }
                }
                return groupedData;
            }
            catch (Exception)
            {
                TagDiffStatic.ReturnValue = 1;
                return [];
            }
        }

        private static FlowRowCompare GenFlowRowCompare(string[] cols, string fileName, string sheet, int lineNum)
        {
            bool isBase = sheet == "base_sheet";
            bool isCompare = sheet == "compare";

            return new FlowRowCompare
            {
                BaseSheet = isBase ? fileName : "",
                ComparedSheet = isCompare ? fileName : "",
                BaseRow = isBase ? lineNum.ToString() : "",
                ComparedRow = isCompare ? lineNum.ToString() : "",

                Label = GetCol(cols, 1),
                Enable = GetCol(cols, 2),
                Job = GetCol(cols, 3),
                Part = GetCol(cols, 4),
                Env = GetCol(cols, 5),
                Opcode = GetCol(cols, 6),
                Parameter = GetCol(cols, 7),
                TName = GetCol(cols, 8),
                TNum = GetCol(cols, 9),
                LoLim = GetCol(cols, 10),
                HiLim = GetCol(cols, 11),
                Scale = GetCol(cols, 12),
                Units = GetCol(cols, 13),
                Format = GetCol(cols, 14),
                BinPass = GetCol(cols, 15),
                BinFail = GetCol(cols, 16),
                SortPass = GetCol(cols, 17),
                SortFail = GetCol(cols, 18),
                Result = GetCol(cols, 19),
                PassAction = GetCol(cols, 20),
                FailAction = GetCol(cols, 21),
                State = GetCol(cols, 22),
                GroupSpecifier = GetCol(cols, 23),
                GroupSense = GetCol(cols, 24),
                GroupCondition = GetCol(cols, 25),
                GroupName = GetCol(cols, 26),
                DeviceSense = GetCol(cols, 27),
                DeviceCondition = GetCol(cols, 28),
                DeviceName = GetCol(cols, 29),
                DebugAssume = GetCol(cols, 30),
                DebugSites = GetCol(cols, 31),
                Comment = GetCol(cols, 37)
            };
        }

        private static string GetCol(string[] columns, int index)
        {
            return index < columns.Length ? columns[index] : "";
        }

        private static bool IsMatchJob(string userJob, string job)
        {
            if (string.IsNullOrEmpty(userJob))
            {
                return true;
            }

            if (string.IsNullOrEmpty(job))
            {
                return true;
            }

            if (job.Contains('!'))
            {
                return !job.Contains(userJob);
            }

            List<string> jobs = [.. job.Split(',')];
            return jobs.Exists(x => x.EqualsIgnoreCase(userJob));
        }
    }
}
