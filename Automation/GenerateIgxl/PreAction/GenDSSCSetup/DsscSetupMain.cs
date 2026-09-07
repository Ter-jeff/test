using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Automation.GenerateIgxl.PreAction.GenDSSCSetup
{
    public static class DsscSetupMain
    {
        private static readonly Dictionary<string, List<DsscSetupSheet>> _dsscsheet =
            new Dictionary<string, List<DsscSetupSheet>>(StringComparer.OrdinalIgnoreCase);

        public static void Initialize()
        {
            _dsscsheet.Clear();
        }

        public static void Save(string sheetName, DsscSetupSheet setup)
        {
            if (!_dsscsheet.TryGetValue(sheetName, out List<DsscSetupSheet> list))
            {
                list = new List<DsscSetupSheet>();
                _dsscsheet[sheetName] = list;
            }

            bool exists = list.Exists(s => s.SetupName.Equals(setup.SetupName, StringComparison.OrdinalIgnoreCase));

            if (!exists)
            {
                list.Add(setup);
            }
        }

        public static Dictionary<string, List<DsscSetupSheet>> GetSheets()
        {
            return _dsscsheet;
        }

        public static List<string> ExportAllSheets(string folderPath)
        {
            List<string> dsscsheets = new List<string>();

            Directory.CreateDirectory(folderPath);

            foreach (KeyValuePair<string, List<DsscSetupSheet>> kv in _dsscsheet)
            {
                string sheetName = kv.Key;
                List<DsscSetupSheet> setups = kv.Value;

                string fileName = $"{sheetName}.txt";
                string filePath = Path.Combine(folderPath, fileName);

                string content = GenerateSingleSheet(setups);

                File.WriteAllText(filePath, content);

                dsscsheets.Add(sheetName);
            }

            return dsscsheets;
        }

        private static string GenerateSingleSheet(List<DsscSetupSheet> setups)
        {
            var sb = new StringBuilder();

            sb.AppendLine(
                "DsscSetup\t" +
                "Pattern\t" +
                "DigSrcEqn\t" +
                "DigSrcReg\t" +
                "DigSrcAssignment\t" +
                "DigSrcPin\t" +
                "DigSrcSampleSize\t" +
                "DigCapPin\t" +
                "DigCapSampleSize\t" +
                "CusStrDigCapData\t" +
                "PatModuleInfo"
            );

            foreach (DsscSetupSheet setup in setups)
            {
                foreach (DsscItem item in setup.Items)
                {
                    DigSrcRegAssi firstReg = item.DigSrcRegAssi != null && item.DigSrcRegAssi.Count > 0
                        ? item.DigSrcRegAssi[0]
                        : null;

                    sb.AppendLine(
                        $"{(item == setup.Items.First() ? setup.SetupName : "")}\t" +
                        $"{item.PatternName}\t" +
                        $"{item.DigSrcEqn}\t" +
                        $"{firstReg?.DigSrcReg ?? ""}\t" +
                        $"{firstReg?.DigSrcAssignment ?? ""}\t" +
                        $"{item.DigSrcPin}\t" +
                        $"{(item.DigSrcSampleSize.Count > 0 ? item.DigSrcSampleSize[0].ToString() : "")}\t" +
                        $"{item.DigCapPin}\t" +
                        $"{item.DigCapSampleSize}\t" +
                        $"{item.CustomDigCapData}\t" +
                        $"{item.PatModuleInfo}"
                    );

                    for (int i = 1; i < (item.DigSrcRegAssi?.Count ?? 0); i++)
                    {
                        DigSrcRegAssi reg = item.DigSrcRegAssi[i];
                        string digSrcSmp = item.DigSrcSampleSize.Count > i ? item.DigSrcSampleSize[i].ToString() : "";

                        sb.AppendLine(
                            $"\t" +
                            $"\t" +
                            $"\t" +
                            $"{reg?.DigSrcReg ?? ""}\t" +
                            $"{reg?.DigSrcAssignment ?? ""}\t" +
                            $"\t" +
                            $"{digSrcSmp}\t" +
                            $"\t" +
                            $"{item.DigCapSampleSize}\t" +
                            $"\t"
                        );
                    }

                }
            }

            return sb.ToString();
        }

    }

    public sealed class DsscSetupSheet
    {
        public string SetupName { get; set; } = string.Empty;

        public List<DsscItem> Items { get; } = new List<DsscItem>();

        // Constructor with name + items
        public DsscSetupSheet(string setupName, List<DsscItem> items)
        {
            SetupName = setupName;
            Items = items ?? new List<DsscItem>();
        }

    }

    public sealed class DsscItem
    {

        private static readonly Regex _regSgmtSize = new Regex(
                @"sgmt(?<id>\d+)_(?<width>\d+)",
                RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public string PatternName { get; set; } = string.Empty;

        // Column: DigSrcEqn (split by '+')
        public string DigSrcEqn { get; set; } = string.Empty;

        // Column: DigSrcAssignment (mapping register name, use 'a:b;a:b;...')
        public List<DigSrcRegAssi> DigSrcRegAssi { get; set; } = new List<DigSrcRegAssi>();

        // Column: DigSrcPin (e.g., CP_BP_JTAG_TDI)
        public string DigSrcPin { get; set; } = string.Empty;

        // Column: DigSrcSampleSize (e.g., "4", "7", "5" for segment widths)
        public List<int> DigSrcSampleSize { get; set; } = new List<int>();

        // Column: DigCapPin and DigCapSampleSize
        public string DigCapPin { get; set; } = string.Empty;
        public int DigCapSampleSize { get; set; }

        // Column: CusStrDigCapData, often holds custom DSSC_OUT or other capture strings
        public string CustomDigCapData { get; set; } = string.Empty;

        public string PatModuleInfo { get; set; } = string.Empty;

        public DsscItem(
                    string patternName,
                    string digSrcEqn,
                    string digSrcAssignment,
                    string digSrcPin,
                    string digSrcSampleSize,
                    string digCapPin,
                    int digCapSampleSize,
                    string customDigCapData,
                    string patModuleInfo)
        {
            PatternName = patternName;
            DigSrcEqn = digSrcEqn ?? string.Empty;
            DigSrcRegAssi = ParseDigSrcAssignment(digSrcAssignment);
            DigSrcPin = digSrcPin ?? string.Empty;
            DigSrcSampleSize = ParseSgmtSizes(digSrcSampleSize);
            DigCapPin = digCapPin ?? string.Empty;
            DigCapSampleSize = digCapSampleSize;
            CustomDigCapData = customDigCapData ?? string.Empty;
            PatModuleInfo = patModuleInfo ?? string.Empty;
        }


        public static List<int> ParseSgmtSizes(string spec)
        {
            List<int> sizes = new List<int>();

            if (string.IsNullOrWhiteSpace(spec))
            {
                return sizes;
            }

            MatchCollection matches = _regSgmtSize.Matches(spec);
            foreach (Match m in matches)
            {
                Group g = m.Groups["width"];
                if (g.Success && int.TryParse(g.Value, out int w))
                {
                    sizes.Add(w);
                }
            }

            return sizes;
        }


        public static List<DigSrcRegAssi> ParseDigSrcAssignment(string assignment)
        {
            List<DigSrcRegAssi> result = new List<DigSrcRegAssi>();

            if (string.IsNullOrWhiteSpace(assignment))
            {
                return result;
            }

            string[] parts = assignment.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string part in parts)
            {
                string[] pair = part.Split(new[] { ':' }, 2, StringSplitOptions.None);

                string reg = pair[0].Trim();
                string assi = pair.Length > 1 ? pair[1].Trim() : string.Empty;

                result.Add(new DigSrcRegAssi(reg, assi));
            }

            return result;
        }

    }

    public sealed class DigSrcRegAssi
    {
        public string DigSrcReg { get; set; } = string.Empty; // signal
        public string DigSrcAssignment { get; set; } = string.Empty; // mapped register

        // Constructor with parameters
        public DigSrcRegAssi(string digSrcReg, string digSrcAssignment)
        {
            DigSrcReg = digSrcReg ?? string.Empty;
            DigSrcAssignment = digSrcAssignment ?? string.Empty;
        }

    }

}
