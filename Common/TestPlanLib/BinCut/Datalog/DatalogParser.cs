using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;

namespace TestPlanLib.BinCut.Datalog
{
    public class DatalogParser
    {
        private const string CurrentJobNameKey = "print: Current job name";
        private const string CurrentJobNameKeyNew = "Current job name";
        private const string Key = "print:";
        private const int MaxCount = 300000;

        private static readonly Dictionary<string, int> _jobStages = new()
        {
            { "CFGTABLE", 0},
            { "CP1_EARLY", 0},
            { "CP1", 0},
            { "CP2", 1},
            { "WLFT", 2},
            { "FT1", 3},
            { "FT2", 4},
            { "FT3", 5},
        };
        public static Job GetDatalogInfo(string dataLogInputFolder)
        {
            var job = new Job(nameof(EnumJob.None));
            if (!string.IsNullOrEmpty(dataLogInputFolder))
            {
                if (!Directory.Exists(dataLogInputFolder))
                {
                    return job;
                }

                List<string> bufFiles = [.. Directory.GetFiles(dataLogInputFolder, "*.txt", SearchOption.TopDirectoryOnly)];
                bufFiles.AddRange([.. Directory.GetFiles(dataLogInputFolder, "*.gz", SearchOption.TopDirectoryOnly)]);
                if (bufFiles.Count > 0)
                {
                    using StreamReader sr = Path.GetExtension(bufFiles[0]).EqualsIgnoreCase(".txt") ? new StreamReader(File.OpenRead(bufFiles[0])) : new StreamReader(new GZipStream(File.OpenRead(bufFiles[0]), CompressionMode.Decompress));
                    GetJobFormLog(sr, out job, out int maxSite, out string prgname);
                }
            }
            return job;
        }

        public static bool GetJobFormLog(StreamReader streamReader, out Job job, out int maxSite, out string programName)
        {
            maxSite = 10;
            job = new Job(nameof(EnumJob.None));
            programName = "";
            int siteCount = 0;
            bool foundSiteNumber = false;
            GetTitleAndProgramInformation(streamReader, out List<string> lines);
            //STEP1. Get max site
            for (int i = 0; i < lines.Count; i++)
            {
                if (string.IsNullOrEmpty(lines[i]))
                {
                    continue;
                }
                //Prog Name:   

                if (lines[i].Contains("Prog Name:"))
                {
                    string[] spt = lines[i].Split([':'], StringSplitOptions.RemoveEmptyEntries);
                    string testpgname = spt[1].Trim().ToUpper();
                    programName = testpgname;
                }
                //Site Number:    
                //0,  1,  2,  3,  4,  5
                else if (lines[i].Contains("Site Number:"))
                {
                    foundSiteNumber = true;
                }
                else if (lines[i].Contains("Job Name:"))
                {
                    string[] spt = lines[i].Split([':'], StringSplitOptions.RemoveEmptyEntries);
                    string testJob = spt[1].Trim().ToUpper();
                    string testJobName = _jobStages.Where(x => testJob.Contains(x.Key)).Select(x => x.Key).DefaultIfEmpty("CP1").First();
                    job = new Job(testJobName);
                }
                else if (lines[i].StartsWithIgnoreCase(CurrentJobNameKey) || lines[i].StartsWithIgnoreCase(CurrentJobNameKeyNew))
                {
                    string testJob = lines[i].Split([',', ':'], StringSplitOptions.RemoveEmptyEntries).Last().Trim();
                    job = new Job(testJob);
                }
                else if (lines[i].Contains("Device#:"))
                {
                    foundSiteNumber = false;
                    if (siteCount != 0)
                    {
                        maxSite = siteCount;
                    }
                }
                else if (foundSiteNumber)
                {
                    siteCount += lines[i].Split([','], StringSplitOptions.RemoveEmptyEntries).Length;
                }
            }
            return true;
        }

        private static void GetTitleAndProgramInformation(StreamReader streamReader, out List<string> lines)
        {
            double cnt = 0;
            lines = [];
            bool isCurrentJobNameKey = false;
            while (!streamReader.EndOfStream)
            {
                string? line = streamReader.ReadLine();
                cnt++;
                if (cnt > MaxCount)
                {
                    return;
                }

                if (line != null && (line.StartsWithIgnoreCase(CurrentJobNameKey) || line.StartsWithIgnoreCase(CurrentJobNameKeyNew)))
                {
                    lines.Add(line);
                    isCurrentJobNameKey = true;
                }
                if (line?.Length != 0)
                {
                    if (isCurrentJobNameKey)
                    {
                        if (!line!.StartsWithIgnoreCase(Key))
                        {
                            break;
                        }
                    }
                    else
                    {
                        lines.Add(line!);
                    }
                }
            }
        }
    }
}
