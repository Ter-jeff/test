using System;

using AutogenCommandLine.CommandLineOptions;
using AutogenCommandLine.Tools.AutoAi;
using AutogenCommandLine.Tools.Autogen;
using AutogenCommandLine.Tools.CAutogen;
using AutogenCommandLine.Tools.CheckScript;
using AutogenCommandLine.Tools.TagDiff;
using AutogenCommandLine.Tools.Validation;

using CommandLine;

using MyCommandLineLib;
using MyCommandLineLib.Enums;

namespace AutogenCommandLine
{
    public class CommandLineFactory
    {
        public static ICommandLineOptions GetOptions(string[] args, ICommandLineParser parser, EntryOptions entryOptions)
        {
            ICommandLineOptions options = new AutogenOptions();
            if (!string.IsNullOrEmpty(entryOptions.InputInfo))
            {
                options = new AutogenOptions();
            }

            if (entryOptions.Entry != null)
            {
                string entry = entryOptions.Entry.ToUpper();
                if (entry == nameof(EnumEntryType.Validation).ToUpper())
                {
                    options = new AutogenOptions();
                }
                else if (entry == nameof(EnumEntryType.Autogen).ToUpper())
                {
                    options = new AutogenOptions();
                }
                else if (entry == nameof(EnumEntryType.BinCutCheck).ToUpper() || entry == nameof(EnumEntryType.BinCutEFuseCheck).ToUpper() || entry == nameof(EnumEntryType.EFuseCheck).ToUpper())
                {
                    options = new CheckScriptOptions();
                }
                else if (entry == nameof(EnumEntryType.BenchLog).ToUpper())
                {
                    options = new BenchLogOptions();
                }
                else if (entry == nameof(EnumEntryType.CAutogen).ToUpper() || entry == nameof(EnumEntryType.CPreChecker).ToUpper())
                {
                    options = new CharAutogenOptions();
                }
                else if (entry == nameof(EnumEntryType.AutoAI).ToUpper())
                {
                    options = new AiAutogenOptions();
                }
                else if (entry == nameof(EnumEntryType.TagDiff).ToUpper())
                {
                    options = new TagDiffOptions();
                }
            }

            parser.ParseArguments(args, options);

            return options;
        }

        public static ICommandLineApplication GetSubCommandLineApplication(IOptions options, EntryOptions entryOptions)
        {
            if (!string.IsNullOrEmpty(entryOptions.InputInfo))
            {
                return new InputInfoCommandLineApp();
            }

            Type optionType = options.GetType();
            ICommandLineApplication cmdApp;
            string entry = entryOptions.Entry.ToUpper();
            if (entry == nameof(EnumEntryType.Validation).ToUpper())
            {
                cmdApp = new ValidationCommandLineApp();
            }
            else if (optionType == typeof(CheckScriptOptions))
            {
                cmdApp = new CheckScriptCommandLineApp();
            }
            else if (entry == nameof(EnumEntryType.CPreChecker).ToUpper())
            {
                cmdApp = new PreCheckerCommandLineApp();
            }
            else if (optionType == typeof(CharAutogenOptions))
            {
                cmdApp = new CAutogenCommandLineApp();
            }
            else if (optionType == typeof(AiAutogenOptions))
            {
                cmdApp = new AutoAiCommandLineApp();
            }
            else if (optionType == typeof(TagDiffOptions))
            {
                cmdApp = new TagDiffCommandLineApp();
            }
            else
            {
                cmdApp = new AutogenCommandLineApp();
            }

            return cmdApp;
        }
    }
}
