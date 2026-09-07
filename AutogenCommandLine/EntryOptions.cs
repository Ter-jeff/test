using System;

using CommandLine;

using MyCommandLineLib;

namespace AutogenCommandLine
{
    public class EntryOptions : IOptions
    {
        [Option('e', "entry", Required = false)]
        public string Entry { get; set; }
        [Option("inputinfo", Required = false)]
        public string InputInfo { get; set; }

        public void GetHelp(string[] args)
        {
            string[] help = [
                "",
                "**** Autogen ****",
                "Command-line arguments:",
                "-?                                        -- get help of tool",
                "",
                "Surround file/folder paths with double quotes if they contain spaces",
                "",
            ];

            foreach (string line in help)
            {
                Console.WriteLine(line);
            }
        }
    }
}
