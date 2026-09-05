using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using AutogenCommandLine.CommandLineOptions;
using AutogenCommandLine.Tools.Autogen;

using Automation;

using CommonLib.Enums;
using CommonLib.ErrorReport;

using LogLib.Static;

using MyCommandLineLib;

namespace AutogenCommandLine.Tools.Validation
{
    public class ValidationCommandLineApp : CommandLineApplicationBase
    {
        public ValidationCommandLineApp()
        {
            ToolName = "Validation";
        }

        public override ICommandLineOptions ValidateInput(ICommandLineOptions options)
        {
            var validationOptions = (AutogenOptions)options;
            if (!string.IsNullOrEmpty(validationOptions.RepoFolderPath))
            {
                validationOptions.Initialize();
            }
            UseStateMachine &= CheckExcelFilePath(validationOptions.TestPlanFile, "-t");

            UseStateMachine &= CheckExcelFilePath(validationOptions.ScghFile, "-s", false);

            UseStateMachine &= CheckCsvFilePath(validationOptions.PatternListCsvFile, "-c", false);

            UseStateMachine &= CheckExcelFilePath(validationOptions.BinCutFile, "-b", false);

            UseStateMachine &= CheckExcelFilePath(validationOptions.PostBinCutFile, "-n", false);

            UseStateMachine &= CheckExcelFilePath(validationOptions.BincutModeSequence, "-k", false);

            UseStateMachine &= CheckCsvFilePath(validationOptions.VoltageTable, "-v", false);

            UseStateMachine &= CheckFolderPathOrTxt(validationOptions.PatternInfoFile, "-g", false);

            UseStateMachine &= CheckOtpFilesPath(validationOptions.OtpFiles, "-f", false);

            UseStateMachine &= CheckYamlFilePath(validationOptions.YamlFile, false);

            UseStateMachine &= CheckFolderPath(validationOptions.PatternFolder, "-a", false);

            UseStateMachine &= CheckFolderPath(validationOptions.TimesetFolder, "-i", false);

            UseStateMachine &= CheckFolderPath(validationOptions.BasLibraryFolder, "-d", false);

            UseStateMachine &= CheckOutputFilePath(validationOptions.OutputDirectory);

            return options;
        }

        public override ICommandLineApplication Execute(ICommandLineOptions options)
        {
            var stopWatch = new Stopwatch();
            stopWatch.Start();
            try
            {
                var args = (AutogenOptions)options;
                AutogenCommandLineApp.PreWorkFlow(args);

                new ValidateTestPlanMain().Run();

                string outputPath = args.OutputDirectory;
                if (outputPath != null)
                {
                    string report = Path.Combine(outputPath, "Error.txt");
                    if (ErrorReportManager.GetErrorList().Count != 0)
                    {
                        IEnumerable<string> lines = ErrorReportManager.GetSortedErrors().Take(20000).Select(x => x.Print());
                        File.AppendAllLines(report, lines);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.StackTrace);
            }
            finally
            {
                stopWatch.Stop();
                Response.Report(string.Format("Total Process Time : " + TimeSpan.FromMilliseconds(stopWatch.ElapsedMilliseconds).ToString(@"hh\:mm\:ss"), EnumMessageLevel.General));
            }
            return this;
        }
    }
}
