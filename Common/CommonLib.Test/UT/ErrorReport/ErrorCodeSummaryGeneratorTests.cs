using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using FileDiffLib;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.ErrorReport
{
    /// <summary>
    /// Generates ErrorCodeSummary.md from the current ErrorCode field definitions and verifies
    /// it still matches the accepted baseline, so documentation drift from the source of truth
    /// is caught in CI instead of silently going stale.
    /// </summary>
    [TestClass]
    public class ErrorCodeSummaryGeneratorTests
    {
        private static readonly string _outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");

        private static readonly string _expectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        private static readonly Regex _placeholderRegex = new(@"\{(\d+)(?::[^}]*)?\}", RegexOptions.Compiled);

        [TestMethod]
        public void GenerateErrorCodeSummaryMarkdownTest()
        {
            string subName = "ErrorCodeSummaryMarkdown";
            string outputPath = Path.Combine(_outputPath, subName);
            string expectPath = Path.Combine(_expectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            (Dictionary<string, (string File, int Line)> fields, Dictionary<string, string> classFiles) = ErrorCodeSourceScanner.Scan();

            Type[] classGroups = [.. typeof(BasicErrorType).Assembly
                .GetTypes()
                .Where(t => t.Namespace == typeof(BasicErrorType).Namespace)
                .Where(t => t != typeof(EfuseCheckCmdLibError))
                .Where(t => t.GetFields(BindingFlags.Public | BindingFlags.Static).Any(f => f.FieldType == typeof(ErrorCode)))
                .OrderBy(t => t.Name, Comparer<string>.Create(string.CompareOrdinal))];

            List<string> sections = [.. classGroups.Select(type => BuildClassSection(type, fields, classFiles))];

            string outputFile = Path.Combine(outputPath, "ErrorCodeSummary.md");
            File.WriteAllText(outputFile, string.Join("\r\n\r\n", sections) + "\r\n", new UTF8Encoding(true));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        private static string BuildClassSection(Type type, Dictionary<string, (string File, int Line)> fields, Dictionary<string, string> classFiles)
        {
            string sourceFile = classFiles.TryGetValue(type.Name, out string found) ? found : $"{type.Name}.cs";

            IEnumerable<FieldInfo> orderedFields = type.GetFields(BindingFlags.Public | BindingFlags.Static)
                .Where(f => f.FieldType == typeof(ErrorCode))
                .OrderBy(f => fields.TryGetValue($"{type.Name}.{f.Name}", out (string File, int Line) loc) ? loc.Line : int.MaxValue);

            IEnumerable<string> entries = orderedFields.Select(fieldInfo => BuildEntry(fieldInfo.Name, (ErrorCode)fieldInfo.GetValue(null)));

            return $"## {type.Name} (`{sourceFile}`)\r\n\r\n" + string.Join("\r\n\r\n", entries);
        }

        private static string BuildEntry(string fieldName, ErrorCode errorCode)
        {
            string template = string.IsNullOrEmpty(errorCode.MessageTemplate) ? "*(empty template)*" : Sanitize(errorCode.MessageTemplate);
            string guidance = Sanitize(errorCode.Guidance);
            int argCount = CountPlaceholders(errorCode.MessageTemplate);
            string behavior = errorCode.EnumErrorBehavior?.ToString() ?? "";
            string target = errorCode.EnumErrorTarget?.ToString() ?? "";

            return $"### {fieldName}\r\n"
                + $"- **FullCode**: {errorCode.FullCode}\r\n"
                + $"- **EnumErrorCategory**: {errorCode.EnumErrorCategory}\r\n"
                + $"- **EnumErrorBehavior**: {behavior}\r\n"
                + $"- **EnumErrorTarget**: {target}\r\n"
                + $"- **Code**: {errorCode.Code}\r\n"
                + $"- **Level**: {errorCode.ErrorLevel}\r\n"
                + $"- **#Args**: {argCount}\r\n"
                + $"- **Template**: {template}\r\n"
                + $"- **Guidance**: {guidance}";
        }

        private static string Sanitize(string value)
        {
            return string.IsNullOrEmpty(value) ? string.Empty : value.Replace("\r\n", " ").Replace("\n", " ").Replace("|", "\\|");
        }

        private static int CountPlaceholders(string template)
        {
            if (string.IsNullOrEmpty(template))
            {
                return 0;
            }

            MatchCollection matches = _placeholderRegex.Matches(template);
            return matches.Select(m => int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture)).DefaultIfEmpty(-1).Max() + 1;
        }
    }
}
