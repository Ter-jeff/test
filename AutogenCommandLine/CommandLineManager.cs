using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

using Automation;

using CommandLine;

using CommonLib.Enums;
using CommonLib.Static;

using LogLib.Static;

using MyCommandLineLib;
using MyCommandLineLib.Enums;


namespace AutogenCommandLine
{
    public class CommandLineManager
    {
        public static bool ParseArgList(string[] args, out ICommandLineOptions options, out EntryOptions entryOptions)
        {
            var settings = new CommandLineParserSettings
            {
                CaseSensitive = false,
                MutuallyExclusive = true,
                IgnoreUnknownArguments = true
            };

            var parser = new CommandLineParser(settings);
            entryOptions = new EntryOptions();
            var helpOptions = new HelpOptions();

            options = null;

            if (args.Length == 0)
            {
                entryOptions.GetHelp(args);
                return false;
            }

            if (!parser.ParseArguments(args, entryOptions))
            {
                entryOptions.GetHelp(args);
                return false;
            }

            if (args.Contains("-?") || args.Contains("?"))
            {
                helpOptions.Entry = entryOptions.Entry;
                helpOptions.GetHelp(args);
                return true;
            }

            if (args.Contains("-r"))
            {
                var assembly = Assembly.GetExecutingAssembly();
                string version = AssemblyProvider.Current.GetFileVersion(FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion);
                var lines = new List<string>
                {
                    "Current version:" + version
                };
                foreach (string line in lines)
                {
                    Console.WriteLine(line);
                }

                return true;
            }

            options = CommandLineFactory.GetOptions(args, parser, entryOptions);
            return true;
        }

        public static void Execute(ICommandLineOptions options, EntryOptions entryOptions)
        {
            ICommandLineApplication cmdApp = CommandLineFactory.GetSubCommandLineApplication(options, entryOptions);
            cmdApp.PreAction(options);
            var assembly = Assembly.GetExecutingAssembly();
            string version = AssemblyProvider.Current.GetFileVersion(FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion);
            Response.Report($"Version: {version}");

            EnumExecuteResult executeResult = EnumExecuteResult.Exception;
            try
            {
                cmdApp.ValidateInput(options);
                cmdApp.Execute(options);
                executeResult = cmdApp.UseStateMachine ? EnumExecuteResult.Done : EnumExecuteResult.Fail;
            }
            catch (CommandLineException ex)
            {
                Response.Report(ex.Message + Environment.NewLine + ex.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
            catch (Exception ex)
            {
                Response.Report("The error message is: unexpected error happened please refer to log file for detail or contact the tool developer.", EnumMessageLevel.Error, 0);
                Response.Report(ex.Message + Environment.NewLine + ex.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
            finally
            {
                Response.Report($"The execution result is: {executeResult}.", EnumMessageLevel.General, 0);
                cmdApp.PostAction(options);
            }
        }
    }
}
