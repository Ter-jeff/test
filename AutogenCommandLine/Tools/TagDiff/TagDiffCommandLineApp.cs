using System.IO;

using AutogenCommandLine.CommandLineOptions;

using Automation;

using MockLib;

using MyCommandLineLib;

using TagDiff.Core;
using TagDiff.Core.Static;

namespace AutogenCommandLine.Tools.TagDiff
{
    public class TagDiffCommandLineApp : CommandLineApplicationBase
    {
        public TagDiffCommandLineApp()
        {
            ToolName = "TagDiff";
        }

        public override ICommandLineOptions ValidateInput(ICommandLineOptions options)
        {
            var tagDiffOptions = (TagDiffOptions)options;

            CheckTestProgramFileFolderPath(tagDiffOptions.AutogenTp);
            CheckTestProgramFileFolderPath(tagDiffOptions.ProdTp);
            CheckOutputFolder(tagDiffOptions.OutDir);

            foreach (string excludeDir in tagDiffOptions.GetExcludeDirs())
            {
                if (!Directory.Exists(excludeDir))
                {
                    string message = $"Exclude directory {excludeDir} not found, -?\" to get help.";
                    throw new CommandLineException(ToolName, message);
                }
            }

            return options;
        }

        public override ICommandLineApplication Execute(ICommandLineOptions options)
        {
            var tagDiffOptions = (TagDiffOptions)options;

            TagDiffStatic.Initialize();

            if (tagDiffOptions.Mock == 1)
            {
                MockService.Mock();
            }

            var calculator = new TagDiffMain(tagDiffOptions.AutogenTp,
                                              tagDiffOptions.ProdTp,
                                              tagDiffOptions.OutDir,
                                              tagDiffOptions.OnlyCsharp,
                                              tagDiffOptions.Job ?? string.Empty,
                                              tagDiffOptions.Mock == 1,
                                              tagDiffOptions.GetExcludeDirs());

            calculator.WorkFlow();

            if (TagDiffStatic.ReturnValue != 0)
            {
                GenerateIgxlMain.ReturnValue = 1;
            }

            return this;
        }
    }
}
