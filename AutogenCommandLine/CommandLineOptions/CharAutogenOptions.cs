using System;
using System.Collections.Generic;

using CommandLine;

using MyCommandLineLib;

namespace AutogenCommandLine.CommandLineOptions
{
    public class CharAutogenOptions : ICommandLineOptions
    {
        [Option('p', "currentproject", Required = true)]
        public string CurrentProjectName { get; set; } = null!;

        [Option("testprogramname", Required = false)]
        public string TestProgramName { get; set; } = null;

        [Option('a', "patternfolder", Required = true)]
        public string PatternFolder { get; set; } = null!;

        [Option("job", Required = true)]
        public string DefaultJob { get; set; } = null!;

        [Option("channelmap", Required = true)]
        public string DefaultChannelMap { get; set; } = null!;

        [Option("charplan", Required = true)]
        public string CharPlan { get; set; } = null!;

        [Option('g', "patterninfofile", Required = true)]
        public string PatternInfoFile { get; set; } = null!;

        [Option('c', "patlistcsvfile", Required = true)]
        public string PatternListCsvFile { get; set; } = null!;

        [Option('o', "outputfolder", Required = true)]
        public string OutputDirectory { get; set; } = null!;

        [Option('b', "basetestprogram", Required = true)]
        public string BaseTestProgram { get; set; } = null!;

        [Option("cslibrarypath", Required = true)]
        public string CsLibraryPath { get; set; } = null!;

        [Option('n', "genpatnotuse", Required = false)]
        public bool GenPatternNotUse { get; set; }

        [Option('r', "gencharnotuse", Required = false)]
        public bool GenCharNotUse
        {
            get; set;
        }

        [Option('m', "mock", Required = false)]
        public int Mock { get; set; } = 0;

        void IOptions.GetHelp(string[] args)
        {
            var help = new List<string>
            {
                "",
                "**** Char-Autogen Tool ****",
                "Command-line arguments:",
                "-p, --currentproject          Current project name(required)              -- Current project name(string)",
                "-a, --patternfolder           Pattern folder(required)                    -- Pattern folder(folder)",
                "--job                         Default job(required)                       -- Default job(string)",
                "--channelmap                  Default channel map(required)               -- Default channel map(string)",
                "--charplan                    Char plan file(required)                    -- Char plan file",
                "-g, --patterninfofile         Pattern info file(required)                 -- Pattern info file(.txt)",
                "-c, --patlistcsvfile          Pattern list CSV file(required)             -- Pattern list CSV file(.csv)",
                "-o, --outputfolder            Output folder(required)                     -- Output folder(folder)",
                "-b, --basetestprogram         Base test program(required)                 -- Base test program(.igxl)",
                "--cslibrarypath               Cs library path(required)                   -- Cs library path(folder)",
                "-n, --genpatnotuse            Generate pattern not use(optional)          -- Generate pattern not use(true/false)",
                "-r, --gencharnotuse           Generate char not use(optional)             -- Generate char not use(true/false)",
                "-m, --mock                    Mock(optional)                              -- Mock(0/1)",
                "",
                "Surround file/folder paths with double quotes if they contain spaces",
                "",
            };

            foreach (string line in help)
            {
                Console.WriteLine(line);
            }
        }

        public string GetOutputFolder()
        {
            return OutputDirectory;
        }
    }
}
