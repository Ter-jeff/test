using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Automation.Static;

namespace Automation.GenerateIgxl.PostAction.VersionTracing
{
    public static class VersionsTable
    {
        public static List<string[]> CreateVersionTable()
        {
            var table = new List<string[]>
            {
                new[] { "Document Type", "Document Filename", "Checksum" }
            };

            foreach (string files in LocalSpecs.SelectedFiles)
            {
                foreach (string filePath in files.Split(','))
                {
                    if (string.IsNullOrEmpty(filePath))
                    {
                        continue;
                    }

                    string documentType = ClassifiedFileByFileName(filePath);
                    string fileName = Path.GetFileName(filePath);
                    string checksum = HashCalculator.CalculateMd5(filePath);
                    table.Add(new[] { documentType, fileName, checksum });
                }
            }
            return table;
        }

        private static string ClassifiedFileByFileName(string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            if (Regex.IsMatch(fileName, "(_EfusePlan|_EFUSE_EXTERNAL).*", RegexOptions.IgnoreCase))
            {
                return "Efuse Test Plan";
            }

            if (Regex.IsMatch(fileName, "_Test.*Plan.*", RegexOptions.IgnoreCase))
            {
                return "Test Plan";
            }

            if (Regex.IsMatch(fileName, "_scgh", RegexOptions.IgnoreCase))
            {
                return "SCGH";
            }

            if (Regex.IsMatch(fileName, "_Pattern.*_.*.csv", RegexOptions.IgnoreCase) && !Regex.IsMatch(fileName, "_TW_V", RegexOptions.IgnoreCase))
            {
                return "Pattern List Csv";
            }

            if (Regex.IsMatch(fileName, @"((\w+)_)?(VoltageTable|Testsetting|VoTa|VolTa)(_\w+)?_(CP|WLFT|FT|FQA|RMA|EMA|T0TxFT)[\d]+(_\w+)?", RegexOptions.IgnoreCase))
            {
                return "Voltage Tables";
            }

            if (Regex.IsMatch(fileName, "(Bin_Cut|Voltage_Binning)", RegexOptions.IgnoreCase))
            {
                return "Bin Cut";
            }

            if (Regex.IsMatch(fileName, "(Post_BinCut|PBC)", RegexOptions.IgnoreCase))
            {
                return "Post Bin Cut";
            }

            if (Regex.IsMatch(fileName, "(bincut_mode_sequence|BMS)", RegexOptions.IgnoreCase))
            {
                return "BinCut Mode Sequence";
            }

            if (Regex.IsMatch(fileName, "(EquationBasedVoltages|EQN)", RegexOptions.IgnoreCase))
            {
                return "Equation Based Voltages";
            }

            if (Regex.IsMatch(fileName, "_TTR_", RegexOptions.IgnoreCase))
            {
                return "HIP TTR Table";
            }

            if (Regex.IsMatch(fileName, "(PowerBinning|PowerScreening)", RegexOptions.IgnoreCase))
            {
                return "PowerBinning";
            }

            if (Regex.IsMatch(fileName, "DRAM", RegexOptions.IgnoreCase))
            {
                return "DRAM Type Table";
            }

            if (Regex.IsMatch(fileName, "FuseCheck", RegexOptions.IgnoreCase))
            {
                return "Fuse Check Table";
            }

            if (Regex.IsMatch(fileName, "output_hashes", RegexOptions.IgnoreCase))
            {
                return "Output Hashes";
            }

            if (Regex.IsMatch(fileName, "ExecInfo", RegexOptions.IgnoreCase))
            {
                return "Exec Info";
            }

            if (Regex.IsMatch(fileName, "_BinOut"))
            {
                return "Bin Out Report";
            }

            return "UnKnown";
        }
    }
}
