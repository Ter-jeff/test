using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CommonLib.Test.UT.ErrorReport
{
    /// <summary>
    /// Accesses one field from each ErrorCodes static class to trigger static constructors
    /// and achieve coverage of all error code definitions.
    /// </summary>
    [TestClass]
    public class ErrorCodesFullCodeTests
    {
        [TestMethod]
        public void AllErrorCodeFields_AcrossAllClasses_HaveUniqueFullCodes()
        {
            // Arrange - reflect over every public static ErrorCode field in every ErrorCodes class
            (Dictionary<string, (string File, int Line)> sourceLocations, _) = ErrorCodeSourceScanner.Scan();

            (string Location, string FullCode, string File, int Line)[] entries = [.. typeof(BinCutErrorType).Assembly
                .GetTypes()
                .Where(t => t.Namespace == typeof(BinCutErrorType).Namespace)
                .SelectMany(t => t.GetFields(BindingFlags.Public | BindingFlags.Static)
                    .Where(f => f.FieldType == typeof(ErrorCode))
                    .Select(f => BuildEntry(t, f, sourceLocations)))];

            // Act
            (string FullCode, (string Location, string File, int Line)[] Locations)[] duplicates = [.. entries
                .GroupBy(e => e.FullCode)
                .Where(g => g.Count() > 1)
                .Select(g => (g.Key, Locations: g.Select(e => (e.Location, e.File, e.Line)).ToArray()))];

            // Assert
            string message = string.Join(System.Environment.NewLine, duplicates.Select(FormatDuplicate));
            Assert.AreEqual(0, duplicates.Length, $"Found {duplicates.Length} duplicate FullCode value(s) across {entries.Length} error codes:{System.Environment.NewLine}{message}");
        }

        private static (string Location, string FullCode, string File, int Line) BuildEntry(System.Type type, FieldInfo fieldInfo, Dictionary<string, (string File, int Line)> sourceLocations)
        {
            string location = $"{type.Name}.{fieldInfo.Name}";
            (string File, int Line) source = sourceLocations.TryGetValue(location, out (string File, int Line) found) ? found : ("<unknown>", 0);
            string fullCode = ((ErrorCode)fieldInfo.GetValue(null)).FullCode;
            return (location, fullCode, source.File, source.Line);
        }

        private static string FormatDuplicate((string FullCode, (string Location, string File, int Line)[] Locations) duplicate)
        {
            string locations = string.Join(System.Environment.NewLine, duplicate.Locations.Select(loc => $"    {loc.Location} ({loc.File}:{loc.Line})"));
            return $"  {duplicate.FullCode}:{System.Environment.NewLine}{locations}";
        }
    }
}
