using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using CommonLib.Extension;

using OfficeOpenXml;

namespace FileDiffLib
{
    public partial class FileComparisonReport
    {
        private const string ReportName = "FileDiffReport";
        private const string Output = "Output";
        private const string Expected = "Expected";

        [GeneratedRegex(@"[#\s\\/:*?""<>|]")]
        private static partial Regex MyRegex();

        private readonly string _name;
        private readonly string _kdiff3;
        private readonly List<string> _excluding;
        private readonly string _oasisKdiff3 = string.Empty;

        public FileComparisonReport(string name, List<string>? excluding = null)
        {
            string workFolder = AppContext.BaseDirectory;
            _kdiff3 = Path.Combine(workFolder, "KDiff3", "kdiff3");
            string? root = Environment.GetEnvironmentVariable("OASISROOT");
            if (!string.IsNullOrEmpty(root))
            {
                _oasisKdiff3 = Path.Combine(root, "KDiff3", "kdiff3.exe");
            }

            _name = name;
            _excluding = excluding ?? [];
        }

        public bool IsFail(string tarDir, string expectedDir, bool isReport, bool isAnyFileType = false)
        {
            ExcelToTxt(tarDir, expectedDir);

            FileComparer fileComparer = isAnyFileType ? GetFileComparer(tarDir, expectedDir, @"^.\w+$") : GetFileComparer(tarDir, expectedDir);
            bool fail = fileComparer.AddItems.Count != 0 || fileComparer.MissingItems.Count != 0 || fileComparer.DiffItems.Count != 0;
            if (fail)
            {
                // Always dump to stdout so CI logs surface the cell-level diffs, regardless of
                // whether the .xlsx report is generated (it is suppressed in Release builds).
                DumpDiffToConsole(fileComparer, _name);
            }
#if RELEASE
            {
                isReport = false;
            }
#endif
            if (isReport && fail)
            {
                GenerateReport(fileComparer, _name);
            }
            return fail;
        }

        private FileComparer GetFileComparer(string tarDir, string expectedDir, string filter = "^.txt$|^.xml$|^.pa$|^.bas$|^.cls$|^.log$|^.igxlProj$|^.atp$|^.json")
        {
            Dictionary<string, string> tar = Directory.Exists(tarDir) ?
                Directory.GetFiles(tarDir, "*.*", SearchOption.AllDirectories)
                .Where(x => Regex.IsMatch(Path.GetExtension(x), filter, RegexOptions.IgnoreCase))
                .ToDictionary(x => x, x => x.Replace(tarDir, "")) : [];
            Dictionary<string, string> exp = Directory.Exists(expectedDir) ?
                Directory.GetFiles(expectedDir, "*.*", SearchOption.AllDirectories)
                .Where(x => Regex.IsMatch(Path.GetExtension(x), filter, RegexOptions.IgnoreCase))
                .ToDictionary(x => x, x => x.Replace(expectedDir, "")) : [];
            var excluding = new List<string>() { "IGLinkManifest.txt", "Autogen.log", "ElapsedTime.json" };
            if (_excluding != null)
            {
                excluding.AddRange(_excluding);
            }

            tar = tar.Where(x => !excluding.Exists(y => y.EqualsIgnoreCase(Path.GetFileName(x.Value))))
                .ToDictionary(x => x.Key, x => x.Value);
            exp = exp.Where(x => !excluding.Exists(y => y.EqualsIgnoreCase(Path.GetFileName(x.Value))))
                .ToDictionary(x => x.Key, x => x.Value);
            FileComparer fileComparer = CompareFiles(tarDir, expectedDir, tar, exp);
            return fileComparer;
        }

        public static FileComparer CompareFiles(string oldDir, string newDir, Dictionary<string, string> oldDic, Dictionary<string, string> newDic)
        {
            var add = oldDic.Values.Except(newDic.Values).ToList();
            var missing = newDic.Values.Except(oldDic.Values).ToList();
            var find = oldDic.Values.Intersect(newDic.Values).ToList();
            var fileComparer = new FileComparer
            {
                AddItems = add,
                MissingItems = missing,
                Output = oldDir,
                Expected = newDir
            };

            var diffBag = new ConcurrentBag<FileDiff>();
            Parallel.ForEach(find, item =>
            {
                if (!TexCompare(oldDir + item, newDir + item))
                {
                    diffBag.Add(new FileDiff { OutputFile = oldDir + item, ExpectedFile = newDir + item, IsSame = false });
                }
            });
            fileComparer.DiffItems.AddRange(diffBag);
            return fileComparer;
        }

        private void ExcelToTxt(string target, string expected)
        {
            var allowed = new HashSet<string>(StringExtensions.IgnoreCase) { ".xls", ".xlsx", ".xlsm" };
            string[] targetFiles = Directory.Exists(target)
                ? [.. Directory.GetFiles(target, "*.*", SearchOption.AllDirectories).Where(file => allowed.Contains(Path.GetExtension(file)))]
                : [];
            string[] expectedFiles = Directory.Exists(expected)
                ? [.. Directory.GetFiles(expected, "*.*", SearchOption.AllDirectories).Where(file => allowed.Contains(Path.GetExtension(file)))]
                : [];

            if (_excluding != null)
            {
                targetFiles = [.. targetFiles.Where(x => !_excluding.Exists(y => Regex.IsMatch(Path.GetFileName(x), y)))];
                expectedFiles = [.. expectedFiles.Where(x => !_excluding.Exists(y => Regex.IsMatch(Path.GetFileName(x), y)))];

            }
            Parallel.ForEach(targetFiles.Concat(expectedFiles), file =>
            {
                string dir = Path.Combine(Path.GetDirectoryName(file) ?? string.Empty, Path.GetFileName(file).Replace(".", "_"));
                Directory.CreateDirectory(dir);
                using var package = new ExcelPackage(new FileInfo(file));
                foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                {
                    string txt = Path.Combine(dir, worksheet.Name + ".txt");
                    worksheet.ExportToTxt(txt);
                }
            });
        }

        public static bool TexCompare(string oldFile, string newFile)
        {
            string[] oldLines = File.ReadAllLines(oldFile);
            string[] newLines = File.ReadAllLines(newFile);
            if (oldLines.Length != newLines.Length)
            {
                return false;
            }

            for (int i = 0; i < oldLines.Length; i++)
            {
                if (oldLines[i] != newLines[i])
                {
                    return false;
                }
            }

            return true;
        }

        private void GenerateReport(FileComparer fileComparer, string name)
        {
            if (fileComparer.AddItems.Count == 0 &&
                fileComparer.MissingItems.Count == 0 &&
                fileComparer.DiffItems.Count == 0)
            {
                return;
            }
            string workFolder = AppContext.BaseDirectory;
            string exportPath = Path.Combine(workFolder, ReportName, name);
            if (Directory.Exists(exportPath))
            {
                try
                {
                    Directory.Delete(exportPath, true);
                }
                catch (Exception)
                {
                    // ignored
                }
            }

            if (!Directory.Exists(exportPath))
            {
                Directory.CreateDirectory(exportPath);
            }

            string reportFile = Path.Combine(exportPath, ReportName + "_" + name + ".xlsx");
            GenDiffReport(reportFile, fileComparer);

            if (File.Exists(reportFile))
            {
                var processStartInfo = new ProcessStartInfo
                {
                    FileName = reportFile,
                    UseShellExecute = true
                };
                Process.Start(processStartInfo);
            }
        }

        private static void DumpDiffToConsole(FileComparer fileComparer, string name)
        {
            const int maxFiles = 5;
            const int maxLinesPerFile = 20;
            Console.WriteLine($"===== FileDiff [{name}] =====");
            foreach (string item in fileComparer.MissingItems)
            {
                Console.WriteLine($"  [MISSING in Output] {item}");
            }
            foreach (string item in fileComparer.AddItems)
            {
                Console.WriteLine($"  [EXTRA in Output]   {item}");
            }
            int fileCount = 0;
            foreach (FileDiff diff in fileComparer.DiffItems)
            {
                if (fileCount++ >= maxFiles)
                {
                    Console.WriteLine($"  ... (+{fileComparer.DiffItems.Count - maxFiles} more diff files suppressed)");
                    break;
                }
                string rel = diff.ExpectedFile.Replace(fileComparer.Expected, "").TrimStart(Path.DirectorySeparatorChar);
                Console.WriteLine($"  --- DIFF: {rel} ---");
                string[] expLines = SafeReadLines(diff.ExpectedFile);
                string[] outLines = SafeReadLines(diff.OutputFile);
                if (expLines.Length != outLines.Length)
                {
                    Console.WriteLine($"    line count differs: expected={expLines.Length} output={outLines.Length}");
                }
                int shown = 0;
                int max = Math.Max(expLines.Length, outLines.Length);
                for (int i = 0; i < max && shown < maxLinesPerFile; i++)
                {
                    string e = i < expLines.Length ? expLines[i] : "<no line>";
                    string o = i < outLines.Length ? outLines[i] : "<no line>";
                    if (e == o)
                    {
                        continue;
                    }

                    int diffAt = FirstDiff(e, o);
                    Console.WriteLine($"    L{i + 1} (lengths exp={e.Length} out={o.Length}, first diff at col {diffAt}):");
                    Console.WriteLine($"      EXP: {WindowAround(e, diffAt, 80)}");
                    Console.WriteLine($"      OUT: {WindowAround(o, diffAt, 80)}");
                    shown++;
                }
                if (shown == maxLinesPerFile)
                {
                    Console.WriteLine($"    ... (additional differing lines suppressed)");
                }
            }
            Console.WriteLine($"===== end FileDiff [{name}] =====");
        }

        private static string[] SafeReadLines(string path)
        {
            try
            {
                return File.ReadAllLines(path);
            }
            catch { return []; }
        }

        private static int FirstDiff(string a, string b)
        {
            int n = Math.Min(a.Length, b.Length);
            for (int i = 0; i < n; i++)
            {
                if (a[i] != b[i])
                {
                    return i;
                }
            }
            return n;
        }

        private static string WindowAround(string s, int pos, int radius)
        {
            int start = Math.Max(0, pos - radius);
            int end = Math.Min(s.Length, pos + radius);
            string head = start > 0 ? "…" : "";
            string tail = end < s.Length ? "…" : "";
            return head + s[start..end] + tail;
        }

        private void GenDiffReport(string diffFile, FileComparer fileComparer)
        {
            var epp = new ExcelPackage(new FileInfo(diffFile));
            ExcelWorksheet sheet = epp.Workbook.Worksheets.InsertSheet(ReportName);
            sheet.Cells[1, 1].Value = Expected;
            sheet.Cells[1, 2].Value = Output;
            sheet.Cells[1, 3].Value = "Same";
            sheet.Cells[1, 4].Value = "Remark";

            int row = 2;
            foreach (string item in fileComparer.MissingItems)
            {
                sheet.Cells[row, 1].Value = item;
                sheet.Cells[row, 4].Value = "Lack of item";
                row++;
            }

            foreach (string item in fileComparer.AddItems)
            {
                sheet.Cells[row, 2].Value = item;
                sheet.Cells[row, 4].Value = "New item";
                row++;
            }

            string? diffFolder = Path.GetDirectoryName(diffFile);
            if (fileComparer.DiffItems.Count > 0 && diffFolder != null)
            {
                if (!Directory.Exists(diffFolder))
                {
                    Directory.CreateDirectory(diffFolder);
                }

                foreach (FileDiff item in fileComparer.DiffItems)
                {
                    string name = item.ExpectedFile.Replace(fileComparer.Expected, "").TrimStart(Path.DirectorySeparatorChar).Replace(Path.DirectorySeparatorChar, '_');

                    string cleanName = MyRegex().Replace(name, "_");
                    string batName = "Diff_" + cleanName + ".bat";
                    string bat = Path.Combine(diffFolder, batName);
                    sheet.Cells[row, 1].Value = item.ExpectedFile.Replace(fileComparer.Expected, "").TrimStart(Path.DirectorySeparatorChar);
                    sheet.Cells[row, 2].Value = item.OutputFile.Replace(fileComparer.Output, "").TrimStart(Path.DirectorySeparatorChar);
                    sheet.Cells[row, 3].Formula = item.IsSame.ToString();
                    sheet.Cells[row, 3].Style.Font.Color.SetColor(item.IsSame ? Color.Green : Color.Red);

                    sheet.Cells[row, 4].Formula = @"=HYPERLINK(""" + batName + @""",""Check the difference"")";
                    sheet.Cells[row, 4].Style.Font.Color.SetColor(Color.Blue);
                    string expectedFile = GetRelativePath(bat, item.ExpectedFile);
                    string outputFile = GetRelativePath(bat, item.OutputFile);
                    string kdiff3 = GetRelativePath(bat, _kdiff3);
                    if (!string.IsNullOrEmpty(_oasisKdiff3))
                    {
                        kdiff3 = _oasisKdiff3;
                    }
                    List<string> lines = GetCommandLines(kdiff3, expectedFile, outputFile);
                    WriteBatFile(bat, lines);

                    row++;
                }
            }

            if (diffFolder != null)
            {
                string bat2 = Path.Combine(diffFolder, "Diff_all.bat");
                string expectedFile = GetRelativePath(bat2, fileComparer.Expected);
                string outputFile = GetRelativePath(bat2, fileComparer.Output);
                string kdiff3 = GetRelativePath(bat2, _kdiff3);
                if (!string.IsNullOrEmpty(_oasisKdiff3))
                {
                    kdiff3 = _oasisKdiff3;
                }
                List<string> lines = GetCommandLines(kdiff3, expectedFile, outputFile);
                WriteBatFile(bat2, lines);
                sheet.Cells[row + 2, 1].Formula = @"=HYPERLINK(""" + expectedFile + @""",""Expected Folder"")";
                sheet.Cells[row + 2, 1].Style.Font.Color.SetColor(Color.Blue);
                sheet.Cells[row + 2, 2].Formula = @"=HYPERLINK(""" + outputFile + @""",""Output Folder"")";
                sheet.Cells[row + 2, 2].Style.Font.Color.SetColor(Color.Blue);
                sheet.Cells[row + 2, 4].Formula = @"=HYPERLINK(""" + Path.GetFileName(bat2) + @""",""Check all"")";
                sheet.Cells[row + 2, 4].Style.Font.Color.SetColor(Color.Blue);
                string bat3 = Path.Combine(diffFolder, "Submit_all.bat");
                string destination = "\"" + Path.Combine(Path.Combine("..", ".."), GetRelativePath(bat2, fileComparer.Expected)) + "\"";
                string source = "\"" + GetRelativePath(bat2, fileComparer.Output) + "\"";
                var lines2 = new List<string>
                {
                    "@echo off",
                    "cd /d %~dp0",
                    "echo Current working directory is: %cd%",
                    // ReSharper disable once StringLiteralTypo
                    "setlocal",
                    "REM Define the source and destination paths",
                    "set source=" + source,
                    "set destination=" + destination,
                    "REM Display the paths for verification",
                    "echo Source folder: %source%",
                    "echo Destination folder: %destination%",
                    "echo.",
                    "REM Check if the destination folder exists",
                    "if exist %destination% (",
                    "    REM Delete the destination folder and all its contents",
                    "    echo Deleting existing destination folder and its contents...",
                    "    rmdir / S / Q %destination%",
                    ")",
                    "REM Create the destination folder",
                    "mkdir %destination%",
                    "REM Copy files and subfolders from source to destination",
                    "echo Copying files from %source% to %destination% ...",
                    // ReSharper disable once StringLiteralTypo
                    "robocopy %source% %destination% /E /COPY:DAT /R:0 /W:0\r\n",
                    "REM Display completion message",
                    "echo Copy complete.",
                    "pause",
                    // ReSharper disable once StringLiteralTypo
                    "endlocal",
                };
                WriteBatFile(bat3, lines2);
                sheet.Cells[row + 3, 4].Formula = @"=HYPERLINK(""" + Path.GetFileName(bat3) + @""",""Submit all"")";
                sheet.Cells[row + 3, 4].Style.Font.Color.SetColor(Color.Blue);
            }

            sheet.Column(1).Width = 80;
            sheet.Column(2).Width = 80;
            sheet.Column(3).Width = 20;
            sheet.Column(4).Width = 80;
            epp.Save();
        }

        private static List<string> GetCommandLines(string kdiff3, string expectedFile, string outputFile)
        {
            return ["@echo off", "cd /d %~dp0", "echo Current working directory is: %cd%", "\"" + kdiff3 + "\" \"" + expectedFile + "\" \"" + outputFile + "\""];
        }

        public static string GetRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath))
            {
                throw new ArgumentNullException(nameof(fromPath));
            }

            if (string.IsNullOrEmpty(toPath))
            {
                throw new ArgumentNullException(nameof(toPath));
            }

            var fromUri = new Uri(fromPath);
            var toUri = new Uri(toPath);

            if (fromUri.Scheme != toUri.Scheme)
            { return toPath; }

            Uri relativeUri = fromUri.MakeRelativeUri(toUri);
            string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            if (toUri.Scheme.EqualsIgnoreCase("file"))
            {
                relativePath = relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            }

            return relativePath;
        }

        private static void WriteBatFile(string outPath, List<string> lines)
        {
            using StreamWriter sw = File.CreateText(outPath);
            foreach (string line in lines)
            {
                sw.WriteLine(line);
            }
        }

        public void GenDiffBats(string batFolder, string outputFolder, string expectedFolder, string oasisRootFolder)
        {
            Directory.CreateDirectory(batFolder);

            string[] outputFiles = Directory.GetFiles(outputFolder, "*.*", SearchOption.TopDirectoryOnly);
            string[] expectedFiles = Directory.GetFiles(expectedFolder, "*.*", SearchOption.TopDirectoryOnly);

            Parallel.ForEach(outputFiles, outputFile =>
            {
                string fileName = Path.GetFileName(outputFile);
                string? expectedFile = expectedFiles.FirstOrDefault(f => Path.GetFileName(f).EqualsIgnoreCase(fileName));
                if (string.IsNullOrEmpty(expectedFile))
                {
                    return;
                }

                string batPath = Path.Combine(batFolder, Path.GetFileNameWithoutExtension(fileName) + ".bat");

                string relativeExpected = GetRelativePath(batPath, expectedFile);
                string relativeOutput = GetRelativePath(batPath, outputFile);
                string kdiff3 = GetRelativePath(batPath, _kdiff3);
                if (!string.IsNullOrEmpty(oasisRootFolder))
                {
                    kdiff3 = Path.Combine(oasisRootFolder, "KDiff3", "kdiff3.exe");
                }
                List<string> lines = GetCommandLines(kdiff3, relativeExpected, relativeOutput);
                WriteBatFile(batPath, lines);
            }
            );
        }
    }
}
