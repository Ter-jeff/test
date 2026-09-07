using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

using AutogenCommandLine.CommandLineOptions;
using AutogenCommandLine.Tools.Autogen;

using Automation.Static;

using CommonLib.Static;

using MyCommandLineLib;

using ProjectConfigLib.ProjectConfig;

namespace AutogenCommandLine.Tools.CAutogen
{
    public class CAutogenCommandLineApp : AutogenCommandLineApp
    {
        public CAutogenCommandLineApp()
        {
            ToolName = "Autogen";
        }

        public override ICommandLineOptions ValidateInput(ICommandLineOptions options)
        {
            var autogenOptions = (CharAutogenOptions)options;

            UseStateMachine &= CheckOutputFolder(autogenOptions.OutputDirectory);

            return options;
        }

        public override ICommandLineApplication Execute(ICommandLineOptions options)
        {
            var args = (CharAutogenOptions)options;
            PreWorkFlow(args);

            PostWorkFlow(args.Mock, args.OutputDirectory);
            return this;
        }

        public static (List<string> tpSheetsList, List<string> scghSheetsList) PreWorkFlow(CharAutogenOptions autogenOptions)
        {
            if (autogenOptions.Mock == 1)
            {
                Mock();
            }

            AllStatic.Clear();
            LocalSpecs.IsUnitTest = autogenOptions.Mock == 1;

            ProjectConfigSingleton.Instance().InitializeProjectConfigSetting();
            ProjectConfigSingleton.Instance().LoadProjectConfig(LocalSpecs.ProjectIniFileName);

            var optionals = new Options(ProjectConfigSingleton.Instance());
            LocalSpecs.Options = optionals;

            LocalSpecs.CurrentProject = autogenOptions.CurrentProjectName;
            LocalSpecs.CurrentJob = autogenOptions.DefaultJob;
            LocalSpecs.TarFolder = autogenOptions.OutputDirectory;
            LocalSpecs.BaseTestProgram = autogenOptions.BaseTestProgram;
            LocalSpecs.TestProgramName = string.IsNullOrEmpty(autogenOptions.TestProgramName) ? autogenOptions.CurrentProjectName : autogenOptions.TestProgramName;

            LocalSpecs.PatternFolder = string.IsNullOrEmpty(autogenOptions.PatternFolder) ? "" : autogenOptions.PatternFolder;

            LocalSpecs.PatternListCsvFileName = string.IsNullOrEmpty(autogenOptions.PatternListCsvFile) ? LocalSpecs.PatternListCsvFileName : autogenOptions.PatternListCsvFile;

            LocalSpecs.CharPlanFileName = autogenOptions.CharPlan;

            LocalSpecs.CsLibraryFolder = string.IsNullOrEmpty(autogenOptions.CsLibraryPath) ? "" : autogenOptions.CsLibraryPath;
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                LocalSpecs.BasLibraryFolder = Path.Combine(Directory.GetCurrentDirectory(), "Settings", "DefaultVbtForCs");
            }

            LocalSpecs.HardIpInfoFileName = string.IsNullOrEmpty(autogenOptions.PatternInfoFile) ? "" : autogenOptions.PatternInfoFile;

            LocalSpecs.DefaultChannelMap = string.IsNullOrEmpty(autogenOptions.DefaultChannelMap) ? "" : autogenOptions.DefaultChannelMap;

            var assembly = Assembly.GetExecutingAssembly();
            string version = AssemblyProvider.Current.GetFileVersion(FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion);
            LocalSpecs.AutogenVer = version;
            List<string> tpSheetsList = [];
            List<string> scghSheetsList = [];
            FolderStructure.CreateFolder();

            return (tpSheetsList, scghSheetsList);
        }
    }
}
