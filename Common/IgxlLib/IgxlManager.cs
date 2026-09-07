using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

using IgxlLib.IgxlSheets;

namespace IgxlLib
{
    public partial class IgxlManager
    {
        private const string TabOrderPrefix = "TAB_ORDER\t";
        private const string TabOrderPropertyFile = "tl_WorkBookProperties_";
        private const string IgLinkClExe = "IGLinkCL.exe";
        private const string IgxlExtension = ".igxl";
        private const string BasExtension = ".bas";
        private const string TextExtension = ".txt";
        private const string OasisRootEnvVar = "OASISROOT";

        [GeneratedRegex("^TAB_ORDER\t", RegexOptions.Compiled)]
        private static partial Regex TabOrderPrefixRegex();

        private static readonly Regex _tabOrderPrefixRegex = TabOrderPrefixRegex();

        private static readonly string _defaultOasisRootFolder = Path.Combine("C:\\", "Program Files (x86)", "teradyne", "oasis");
        private readonly string _oasisRootFolder;

        public IgxlManager()
        {
            _oasisRootFolder = GetOasisRootFolder();
        }

        private static string GetOasisRootFolder()
        {
            string? oasisRoot = Environment.GetEnvironmentVariable(OasisRootEnvVar);
            return string.IsNullOrEmpty(oasisRoot) ? _defaultOasisRootFolder : oasisRoot;
        }

        public static void AddIgxlSheets(string igxl, List<IIgxlSheet> igxlSheets, string tempFolder)
        {
            var lines = new List<string>();
            using var fs = new FileStream(igxl, FileMode.OpenOrCreate);
            using var archive = new ZipArchive(fs, ZipArchiveMode.Update);
            ExtractTabOrderFromArchive(archive, lines);
            UpdateTabOrderIfNeeded(igxlSheets, lines);
            UpdatePropertyFile(archive, lines);
            UpdateIgxlSheets(archive, igxlSheets, tempFolder);
        }

        private static void ExtractTabOrderFromArchive(ZipArchive zipArchive, List<string> lines)
        {
            foreach (ZipArchiveEntry entry in zipArchive.Entries)
            {
                string sheetName = Path.GetFileNameWithoutExtension(entry.FullName);
                if (sheetName.EqualsIgnoreCase(TabOrderPropertyFile))
                {
                    using Stream stream = entry.Open();
                    using var sr = new StreamReader(stream);
                    while (!sr.EndOfStream)
                    {
                        lines.Add(sr.ReadLine()!);
                    }
                    return;
                }
            }
        }

        private static void UpdateTabOrderIfNeeded(List<IIgxlSheet> igxlSheets, List<string> lines)
        {
            if (lines.Count == 0 || !lines.First().StartsWithIgnoreCase(TabOrderPrefix))
            {
                return;
            }

            string line = lines.First();
            line = _tabOrderPrefixRegex.Replace(line, "");
            var arr = line.Split(';').Where(x => !string.IsNullOrEmpty(x)).ToList();
            var sheetNames = arr.Select(x => x.Split(':').First()).ToList();
            sheetNames.AddRange(igxlSheets.Where(x => x is not BasFile).Select(x => x.Name));
            sheetNames = [.. sheetNames.OrderBy(x => x)];
            string addText = string.Join(";", sheetNames.Select((x, idx) => x + ":" + (idx + 1)));
            string newLine = TabOrderPrefix + addText + ";";
            lines[0] = newLine;
        }

        private static void UpdatePropertyFile(ZipArchive zipArchive, List<string> lines)
        {
            string propEntryName = TabOrderPropertyFile + TextExtension;
            zipArchive.GetEntry(propEntryName)?.Delete();
            ZipArchiveEntry newPropEntry = zipArchive.CreateEntry(propEntryName);
            using var writer = new StreamWriter(newPropEntry.Open());
            writer.Write(string.Join(Environment.NewLine, lines));
        }

        private static void UpdateIgxlSheets(ZipArchive zipArchive, List<IIgxlSheet> igxlSheets, string tempFolder)
        {
            foreach (IIgxlSheet igxlSheet in igxlSheets)
            {
                string extension = igxlSheet is BasFile ? BasExtension : TextExtension;
                string fileName = igxlSheet.Name + extension;
                string zipEntryName = fileName.Replace(" ", "%20");

                string tempPath = Path.Combine(tempFolder, fileName);
                igxlSheet.Write(tempPath);

                zipArchive.GetEntry(zipEntryName)?.Delete();
                ZipArchiveEntry sheetEntry = zipArchive.CreateEntry(zipEntryName);
                using var writer = new StreamWriter(sheetEntry.Open());

                if (File.Exists(tempPath))
                {
                    writer.Write(File.ReadAllText(tempPath));
                }
            }
        }

        public static bool ExportWorkBook(string testProgramFile, string exportFolder)
        {
            RecreateExportDirectory(exportFolder);

            if (Path.GetExtension(testProgramFile).EqualsIgnoreCase(IgxlExtension))
            {
                ZipFile.ExtractToDirectory(testProgramFile, exportFolder);
                return true;
            }

            return false;
        }

        private static void RecreateExportDirectory(string exportFolder)
        {
            if (Directory.Exists(exportFolder))
            {
                Directory.Delete(exportFolder, true);
            }
            Directory.CreateDirectory(exportFolder);
        }

        public void GenTestProgramBySubProgram(string igLinkProject, string outputFile, string switchStr, string subProgramName)
        {
            string exe = Path.Combine(_oasisRootFolder, IgLinkClExe);
            if (!File.Exists(exe))
            {
                return;
            }

            string option = $"-i \"{igLinkProject}\" -s \"{subProgramName}\"{switchStr}\"{outputFile}\"";
            RunProcess(exe, option);
        }

        private static bool RunProcess(string fileName, string arguments = "", bool captureOutput = false)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    FileName = fileName,
                    Arguments = arguments,
                    RedirectStandardOutput = captureOutput,
                    RedirectStandardError = captureOutput
                };

                using var process = new Process();
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();
                return process.ExitCode == 0;
            }
            catch (Exception)
            {
                // Log or handle exception as appropriate for the application
                return false;
            }
        }
    }
}
