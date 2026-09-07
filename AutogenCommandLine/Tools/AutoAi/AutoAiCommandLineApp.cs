using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

using AutogenCommandLine.CommandLineOptions;

using Automation.Static;

using Cautogen;

using CommonLib.Static;

using LogLib.Static;

using MyCommandLineLib;

namespace AutogenCommandLine.Tools.AutoAi
{
    public class AutoAiCommandLineApp : CommandLineApplicationBase
    {
        public AutoAiCommandLineApp()
        {
            ToolName = "Autogen";
        }

        public override ICommandLineOptions ValidateInput(ICommandLineOptions options)
        {
            var autogenOptions = (AiAutogenOptions)options;

            UseStateMachine &= CheckOutputFolder(autogenOptions.OutputDirectory);

            return options;
        }

        public override ICommandLineApplication Execute(ICommandLineOptions options)
        {
            var args = (AiAutogenOptions)options;
            PreWorkFlow(args);

            Response.Report("Start AI-Autogen ...");
            new AiAutogenMain().Execute();
            Response.Report(" AI-Autogen done.");

            return this;
        }

        public static (List<string> tpSheetsList, List<string> scghSheetsList) PreWorkFlow(AiAutogenOptions autogenOptions)
        {
            AllStatic.Clear();

            LocalSpecs.CurrentProject = autogenOptions.CurrentProjectName;
            LocalSpecs.CurrentJob = autogenOptions.DefaultJob;
            LocalSpecs.TarFolder = autogenOptions.OutputDirectory;
            LocalSpecs.BaseTestProgram = autogenOptions.BaseTestProgram;
            LocalSpecs.TestProgramName = string.IsNullOrEmpty(autogenOptions.TestProgramName) ? autogenOptions.CurrentProjectName : autogenOptions.TestProgramName;

            LocalSpecs.PatternFolder = string.IsNullOrEmpty(autogenOptions.PatternFolder) ? "" : autogenOptions.PatternFolder;
            LocalSpecs.DefaultChannelMap = string.IsNullOrEmpty(autogenOptions.DefaultChannelMap) ? "" : autogenOptions.DefaultChannelMap;

            LocalSpecs.CsLibraryFolder = string.IsNullOrEmpty(autogenOptions.CsLibraryPath) ? "" : autogenOptions.CsLibraryPath;
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                LocalSpecs.BasLibraryFolder = Path.Combine(Directory.GetCurrentDirectory(), "Settings", "DefaultVbtForCs");
            }

            LocalSpecs.CharPlanFileName = autogenOptions.CharPlan;

            var assembly = Assembly.GetExecutingAssembly();
            string version = AssemblyProvider.Current.GetFileVersion(FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion);
            LocalSpecs.AutogenVer = version;
            List<string> tpSheetsList = [];
            List<string> scghSheetsList = [];

            return (tpSheetsList, scghSheetsList);
        }
    }
}
