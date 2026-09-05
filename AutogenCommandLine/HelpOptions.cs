using System;

using AutogenCommandLine.CommandLineOptions;

using CommandLine;

using MyCommandLineLib;
using MyCommandLineLib.Enums;

namespace AutogenCommandLine
{
    public class HelpOptions : IOptions
    {
        public string Entry { get; set; }

        [HelpOption("?")]
        public void GetHelp(string[] args)
        {
            IOptions opts = ResolveHelpOptions();
            opts.GetHelp(args);
        }

        private IOptions ResolveHelpOptions()
        {
            if (Entry == null)
            {
                return new AutogenOptions();
            }

            if (string.Equals(Entry, nameof(EnumEntryType.Autogen), StringComparison.CurrentCultureIgnoreCase) ||
                string.Equals(Entry, nameof(EnumEntryType.Validation), StringComparison.CurrentCultureIgnoreCase))
            {
                return new AutogenOptions();
            }

            if (string.Equals(Entry, nameof(EnumEntryType.BinCutCheck), StringComparison.CurrentCultureIgnoreCase) ||
                string.Equals(Entry, nameof(EnumEntryType.BinCutEFuseCheck), StringComparison.CurrentCultureIgnoreCase) ||
                string.Equals(Entry, nameof(EnumEntryType.EFuseCheck), StringComparison.CurrentCultureIgnoreCase))
            {
                return new CheckScriptOptions();
            }

            if (string.Equals(Entry, nameof(EnumEntryType.BenchLog), StringComparison.CurrentCultureIgnoreCase))
            {
                return new BenchLogOptions();
            }

            if (string.Equals(Entry, nameof(EnumEntryType.CAutogen), StringComparison.CurrentCultureIgnoreCase) ||
                string.Equals(Entry, nameof(EnumEntryType.CPreChecker), StringComparison.CurrentCultureIgnoreCase))
            {
                return new CharAutogenOptions();
            }

            if (string.Equals(Entry, nameof(EnumEntryType.AutoAI), StringComparison.CurrentCultureIgnoreCase))
            {
                return new AiAutogenOptions();
            }

            return new AutogenOptions();
        }
    }
}
