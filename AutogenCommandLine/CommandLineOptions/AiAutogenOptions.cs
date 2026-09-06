using System;
using System.Collections.Generic;

using CommandLine;

using MyCommandLineLib;

namespace AutogenCommandLine.CommandLineOptions
{
    public class AiAutogenOptions : ICommandLineOptions
    {
        [Option('p', "currentproject", Required = true)]
        public string CurrentProjectName { get; set; } = null!;

        [Option("testprogramname", Required = false)]
        public string TestProgramName { get; set; } = null!;

        [Option('a', "patternfolder", Required = true)]
        public string PatternFolder { get; set; } = null!;

        [Option("channelmap", Required = true)]
        public string DefaultChannelMap { get; set; } = null!;

        [Option('b', "basetestprogram", Required = true)]
        public string BaseTestProgram { get; set; } = null!;

        [Option("cslibrarypath", Required = true)]
        public string CsLibraryPath { get; set; } = null!;

        [Option('o', "outputfolder", Required = true)]
        public string OutputDirectory { get; set; } = null!;

        [Option("job", Required = true)]
        public string DefaultJob { get; set; } = null!;

        [Option("charplan", Required = true)]
        public string CharPlan { get; set; } = null!;

        void IOptions.GetHelp(string[] args)
        {
            var help = new List<string>
            {
                "",
                "**** AI-Autogen Tool ****",
                "Command-line arguments:",
                "-p, --currentproject          Current project name(required)              -- Current project name(string)",
                "-a, --patternfolder           Pattern folder(required)                    -- Pattern folder(folder)",
                "--channelmap                  Default channel map(required)               -- Default channel map(string)",
                "-b, --basetestprogram         Base test program(required)                 -- Base test program(.igxl)",
                "--cslibrarypath               Cs library path(required)                   -- Cs library path(folder)",
                "-o, --outputfolder            Output folder(required)                     -- Output folder(folder)",
                "--job                         Default job(required)                       -- Default job(string)",
                "--charplan                    Char plan file(required)                    -- Char plan file",
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
