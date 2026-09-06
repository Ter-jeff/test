using System;
using System.Collections.Generic;

using CommandLine;

using MyCommandLineLib;

namespace AutogenCommandLine.CommandLineOptions
{
    public class TagDiffOptions : ICommandLineOptions
    {
        [Option('a', "autogen-tp", Required = true, HelpText = "Path to autogen test program .igxl or .xls file")]
        public string AutogenTp { get; set; } = null!;

        [Option('p', "prod-tp", Required = true, HelpText = "Path to production test program .igxl or .xls file")]
        public string ProdTp { get; set; } = null!;

        [Option('o', "out-dir", Required = true, HelpText = "Output directory")]
        public string OutDir { get; set; } = null!;

        [Option('j', "job", Required = false, HelpText = "Optional job name")]
        public string Job { get; set; } = null!;

        [Option('c', "only-csharp", DefaultValue = false, HelpText = "Only run C# instance comparison")]
        public bool OnlyCsharp { get; set; }

        [Option('m', "mock", Required = false)]
        public int Mock { get; set; }

        [Option("exclude_dirs", Required = false, HelpText = "Optional comma-separated list of directories to exclude from comparison")]
        public string ExcludeDirs { get; set; } = null!;

        public IEnumerable<string> GetExcludeDirs()
        {
            return string.IsNullOrEmpty(ExcludeDirs)
                ? []
                : ExcludeDirs.Split(',');
        }

        public string GetOutputFolder()
        {
            return OutDir;
        }

        public void GetHelp(string[] args)
        {
            var help = new List<string>
            {
                "",
                "usage: AutogenCommandLine.exe -e TagDiff -a AUTOGEN_TP -p PROD_TP -o OUT_DIR [-c]",
                "",
                "required arguments:",
                "  -a AUTOGEN_TP, --autogen-tp AUTOGEN_TP",
                "        Path to autogen test program .igxl or .xls file",
                "  -p PROD_TP, --prod-tp PROD_TP",
                "        Path to production test program .igxl or .xls file",
                "  -o OUT_DIR, --out-dir OUT_DIR",
                "        Output directory",
                "",
                "optional arguments:",
                "  -h, --help",
                "        Show this help message and exit",
                "  -c, --only-csharp",
                "        Only run C# instance comparison",
                "  -j, --job",
                "        Optional job name",
                "  --exclude_dirs",
                "        Optional list of directories to exclude from comparison",
            };

            foreach (string line in help)
            {
                Console.WriteLine(line);
            }
        }
    }
}
