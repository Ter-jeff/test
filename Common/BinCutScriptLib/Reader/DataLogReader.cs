using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib;
using TestPlanLib.BinCut.Datalog;

namespace BinCutScriptLib.Reader
{
    public class DataLogReader
    {
        public static bool ReadJob(string dataLogFile, out int maxSite, out string programName, out Job job, Action<string, Color> richTextBoxAppend)
        {
            using StreamReader sr = GetStreamReader(dataLogFile);
            if (!DatalogParser.GetJobFormLog(sr, out job, out maxSite, out programName))
            {
                string errorMessage = "Can't found Job Name over 300000 lines !!!";
                richTextBoxAppend?.Invoke(errorMessage, Color.Red);

                return false;
            }

            if (job.JobType == EnumJob.None)
            {
                string errorMessage = "Job name is incorrect !!!";
                richTextBoxAppend?.Invoke(errorMessage, Color.Red);

                return false;
            }
            return true;
        }

        public static StreamReader GetStreamReader(string dataLogFile)
        {
            string extension = Path.GetExtension(dataLogFile);
            return extension != null && extension.EqualsIgnoreCase(".txt") ?
                new StreamReader(File.OpenRead(dataLogFile)) :
                new StreamReader(new GZipStream(File.OpenRead(dataLogFile), CompressionMode.Decompress));
        }

        public static bool CheckDatalog(string file, string programFileName, out Job job, Action<string, Color> richTextBoxAppend)
        {
            if (!ReadJob(file, out int _, out string programName, out job, richTextBoxAppend))
            {
                return true;
            }

            BinCutDatalogConfigReader.ReadBinCutConfig(file, ref job);

            if (!string.IsNullOrEmpty(programFileName) && !string.IsNullOrEmpty(programName))
            {
                string name = Path.GetFileName(programFileName);
                if (!name.EqualsIgnoreCase(programName))
                {
                    richTextBoxAppend("The program name of datalog is different with selected test program file !!!", Color.Red);
                    richTextBoxAppend($"Program name in data log is {programName} !!!", Color.Red);
                    richTextBoxAppend($"Program name is {name} !!!", Color.Red);
                }
            }
            return false;
        }

        public static bool CheckDatalog(string dataLogInputFolder, string programFileName, out Job job, Action<string, Color> richTextBoxAppend, out List<string> files, out bool csFlag)
        {
            csFlag = false;
            job = new Job(nameof(EnumJob.None));
            files = [.. Directory.GetFiles(dataLogInputFolder, "*.txt", SearchOption.TopDirectoryOnly)];
            files.AddRange([.. Directory.GetFiles(dataLogInputFolder, "*.gz", SearchOption.TopDirectoryOnly)]);
            if (files.Count == 0)
            {
                richTextBoxAppend("Can't find any log to parse, program terminated.", Color.Red);
                return false;
            }
            richTextBoxAppend($"Found {files.Count} log(s): ", Color.Blue);

            CheckDirectory(dataLogInputFolder, richTextBoxAppend);

            if (!ReadJob(files.First(), out int _, out string programName, out job, richTextBoxAppend))
            {
                return true;
            }

            csFlag = CheckDatalogFormatNew(files.First());

            if (csFlag)
            {
                BinCutDatalogConfigReader.ReadBinCutConfigCs(files.First(), ref job, programFileName);
            }
            else
            {
                BinCutDatalogConfigReader.ReadBinCutConfig(files.First(), ref job);
            }

            if (!string.IsNullOrEmpty(programFileName) && !string.IsNullOrEmpty(programName))
            {
                string name = Path.GetFileName(programFileName);
                if (!name.EqualsIgnoreCase(programName))
                {
                    richTextBoxAppend("The program name of datalog is different with selected test program file !!!", Color.Red);
                    richTextBoxAppend($"Program name in data log is {programName} !!!", Color.Red);
                    richTextBoxAppend($"Program name is {name} !!!", Color.Red);
                }
            }
            return false;
        }

        public static bool CheckDatalog(string dataLogInputFolder, out bool csFlag)
        {
            csFlag = false;
            List<string> files = [.. Directory.GetFiles(dataLogInputFolder, "*.txt", SearchOption.TopDirectoryOnly)];
            files.AddRange([.. Directory.GetFiles(dataLogInputFolder, "*.gz", SearchOption.TopDirectoryOnly)]);
            if (files.Count == 0)
            {
                return false;
            }

            csFlag = CheckDatalogFormatNew(files.First());

            return true;
        }

        protected static void CheckDirectory(string dataLogInputFolder, Action<string, Color> richTextBoxAppend)
        {
            if (!Directory.Exists(dataLogInputFolder))
            {
                richTextBoxAppend($"Can't find {dataLogInputFolder} !!!", Color.Red);
            }
        }

        public static bool CheckDatalogFormat(string dataLogInputFolder, string programFileName, Action<string, Color> richTextBoxAppend, out List<string> files)
        {
            files = [.. Directory.GetFiles(dataLogInputFolder, "*.txt", SearchOption.TopDirectoryOnly)];
            files.AddRange([.. Directory.GetFiles(dataLogInputFolder, "*.gz", SearchOption.TopDirectoryOnly)]);
            if (files.Count == 0)
            {
                richTextBoxAppend("Can't find any log to parse, program terminated.", Color.Red);
            }
            richTextBoxAppend($"Found {files.Count} log(s): ", Color.Blue);
            if (files.Count != 0)
            {
                using var sr = new StreamReader(files.First());
                string? line;
                while ((line = sr.ReadLine()) != null)
                {
                    Match csharpMatchFlag = Reg.RegexcsharpMatchFlag.Match(line);
                    Match vbtMatchFlag = Reg.RegexvbtMatchFlag.Match(line);
                    if (csharpMatchFlag.Success)
                    {
                        return true;
                    }
                    if (vbtMatchFlag.Success)
                    {
                        return false;
                    }
                    if (line.Length >= 1000)
                    {
                        return false;
                    }

                }
            }
            return true;
        }

        public static bool CheckDatalogFormatNew(string dataLogFile)
        {
            using StreamReader sr = GetStreamReader(dataLogFile);
            string? line;
            while ((line = sr.ReadLine()) != null)
            {
                Match csharpMatchFlag = Reg.RegexcsharpMatchFlag.Match(line);
                Match vbtMatchFlag = Reg.RegexvbtMatchFlag.Match(line);
                if (csharpMatchFlag.Success)
                {
                    return true;
                }
                if (vbtMatchFlag.Success)
                {
                    return false;
                }
                if (line.Length >= 1000)
                {
                    return false;
                }

            }
            return true;
        }

    }
}
