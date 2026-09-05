using System;
using System.Collections.Generic;

using CommandLine;

using MyCommandLineLib;

namespace AutogenCommandLine.CommandLineOptions
{
    public class CheckScriptOptions : ICommandLineOptions
    {
        [Option('e', "", Required = false, HelpText = "Type")]
        public string Type { get; set; } = null!;

        [Option('p', "program", Required = false, HelpText = "Input file : Program")]
        public string Prog { get; set; } = null!;

        [Option('f', "logfolder", Required = true, HelpText = "Data log file")]
        public string Datalog { get; set; } = null!;

        [Option('c', "config", Required = false, HelpText = "Input file : Efuse Config")]
        public string Config { get; set; } = null!;
        [Option('b', "bitdef", Required = false, HelpText = "Input file : BitDef Table")]
        public string Bdf { get; set; } = null!;

        [Option('o', "otpfilesfolder", HelpText = "Output files folder")]
        public string OutputFolder { get; set; } = null!;

        public string GetOutputFolder()
        {
            return OutputFolder;
        }

        void IOptions.GetHelp(string[] args)
        {
            var help = new List<string>
            {
                "",
                "**** CheckScript Tool ****",
                "Command-line arguments:",
                "-e                            Check type(optional)                        -- Type",
                "-p, --program                 Program(optional)                           -- Input file : Program",
                "-f, --logfolder               Data log file(required)                     -- Data log file",
                "-c, --config                  Efuse config(optional)                      -- Input file : Efuse Config",
                "-b, --bitdef                  BitDef table(optional)                      -- Input file : BitDef Table",
                "-o, --otpfilesfolder          Output files folder(optional)               -- Output files folder",
                "",
                "Surround file/folder paths with double quotes if they contain spaces",
                "",
            };

            foreach (string line in help)
            {
                Console.WriteLine(line);
            }
        }
    }
}
