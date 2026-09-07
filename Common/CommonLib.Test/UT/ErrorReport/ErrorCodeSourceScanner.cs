using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace CommonLib.Test.UT.ErrorReport
{
    /// <summary>
    /// Scans the ErrorCodes source files to recover, per field, the declaring file name
    /// and line number - information reflection alone cannot provide.
    /// </summary>
    internal static class ErrorCodeSourceScanner
    {
        private static readonly Regex _classRegex = new(@"class\s+(\w+)", RegexOptions.Compiled);

        private static readonly Regex _fieldRegex = new(@"public\s+static\s+readonly\s+ErrorCode\s+(\w+)\s*=", RegexOptions.Compiled);

        public static string ResolveErrorCodesDirectory([CallerFilePath] string thisFilePath = "")
        {
            return Path.GetFullPath(Path.Combine(Path.GetDirectoryName(thisFilePath) ?? ".", "..", "..", "..", "CommonLib", "ErrorReport", "ErrorCodes"));
        }

        public static (Dictionary<string, (string File, int Line)> Fields, Dictionary<string, string> ClassFiles) Scan()
        {
            var fields = new Dictionary<string, (string File, int Line)>();
            var classFiles = new Dictionary<string, string>();
            string errorCodesDir = ResolveErrorCodesDirectory();

            if (!Directory.Exists(errorCodesDir))
            {
                return (fields, classFiles);
            }

            foreach (string file in Directory.GetFiles(errorCodesDir, "*.cs"))
            {
                ScanFile(file, fields, classFiles);
            }

            return (fields, classFiles);
        }

        private static void ScanFile(string file, Dictionary<string, (string File, int Line)> fields, Dictionary<string, string> classFiles)
        {
            string[] lines = File.ReadAllLines(file);
            string currentClass = string.Empty;
            string fileName = Path.GetFileName(file);

            for (int i = 0; i < lines.Length; i++)
            {
                Match classMatch = _classRegex.Match(lines[i]);
                if (classMatch.Success)
                {
                    currentClass = classMatch.Groups[1].Value;
                    classFiles[currentClass] = fileName;
                    continue;
                }

                Match fieldMatch = _fieldRegex.Match(lines[i]);
                if (fieldMatch.Success && currentClass.Length > 0)
                {
                    string key = $"{currentClass}.{fieldMatch.Groups[1].Value}";
                    fields[key] = (fileName, i + 1);
                }
            }
        }
    }
}
