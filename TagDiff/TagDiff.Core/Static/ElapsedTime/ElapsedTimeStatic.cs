using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using Newtonsoft.Json;

using TagDiff.Core.Output;

namespace TagDiff.Core.Static.ElapsedTime
{
    internal static class ElapsedTimeStatic
    {
        private static readonly object _lock = new();
        public static List<ElapsedTimeItem> ElapsedTimeItems { get; set; } = [];

        public static void WriteToJson(string outputPath, List<ElapsedTimeItem> elapsedTimeItems, List<CompareReport> compareReports)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("outputPath must be a valid directory path.", nameof(outputPath));
            }

            if (elapsedTimeItems == null || elapsedTimeItems.Count == 0 || compareReports == null || compareReports.Count == 0)
            {
                return;
            }

            List<ElapsedTimeResult> results = GetElapsedTimeResults(elapsedTimeItems, compareReports);

            string outputFile = Path.Combine(outputPath, "ElapsedTime.json");
            string json = JsonConvert.SerializeObject(results, Formatting.Indented);
            File.WriteAllText(outputFile, json);
        }

        private static List<ElapsedTimeResult> GetElapsedTimeResults(List<ElapsedTimeItem> elapsedTimeItems, List<CompareReport> compareReports)
        {
            // Build a case-insensitive lookup for sheet name -> timespan (O(n))
            var timeLookup = new Dictionary<string, TimeSpan>(StringExtensions.IgnoreCase);
            if (elapsedTimeItems != null)
            {
                foreach (ElapsedTimeItem item in elapsedTimeItems)
                {
                    if (string.IsNullOrEmpty(item?.SheetName))
                    {
                        continue;
                    }

                    if (!timeLookup.ContainsKey(item.SheetName))
                    {
                        timeLookup[item.SheetName] = item.Time;
                    }
                }
            }

            // Group compare reports by Category directly and sum ticks for each group (O(m))
            var results = compareReports.Where(cr => cr != null).GroupBy(cr => cr.Category ?? string.Empty).Select(group =>
                {
                    long totalTicks = 0L;
                    foreach (CompareReport cr in group)
                    {
                        if (string.IsNullOrEmpty(cr?.SheetName))
                        {
                            continue;
                        }

                        if (timeLookup.TryGetValue(cr.SheetName, out TimeSpan ts))
                        {
                            totalTicks += ts.Ticks;
                        }
                    }

                    var sum = new TimeSpan(totalTicks);
                    return new ElapsedTimeResult(group.Key, sum.TotalSeconds.ToString(CultureInfo.InvariantCulture));
                })
                .ToList();

            return results;
        }

        public static void Add(ElapsedTimeItem elapsedTimeItem)
        {
            lock (_lock)
            {
                ElapsedTimeItems.Add(elapsedTimeItem);
            }
        }

        public static void Clear()
        {
            lock (_lock)
            {
                if (ElapsedTimeItems != null)
                {
                    ElapsedTimeItems.Clear();
                }
                else
                {
                    ElapsedTimeItems = [];
                }
            }
        }
    }
}
