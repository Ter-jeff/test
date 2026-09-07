using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;

using Automation.Static;

using BinCutScriptLib;
using BinCutScriptLib.Reader;
using BinCutScriptLib.Static;

using LogLib.Static;

using NLog;

using ProjectConfigLib.ProjectConfig;

using TestPlanLib;

namespace AutogenCommandLine.Tools.CheckScript
{
    public class BinCut
    {
        private static string _binCutSpec;
        private static string _postBinCutSpec;
        private readonly string _inputFile;
        private readonly string _testProgram;
        private readonly string _outputFolder;
        private static string _testPlan;
        private static string _idsDistribution;

        public BinCut(string inputFile, string inPutProg, string outPutPath)
        {
            _binCutSpec = "";
            _postBinCutSpec = "";
            _inputFile = inputFile;
            _testProgram = inPutProg;
            _outputFolder = outPutPath;
            _testPlan = "";
            _idsDistribution = "";
        }

        public void RunBinCut()
        {
            if (!string.IsNullOrEmpty(_inputFile) && !string.IsNullOrEmpty(_testProgram))
            {
                Response.Report("Start BinCut Check Script...");
                DataLogReader.CheckDatalog(_inputFile, _testProgram, out Job job, AppendText, out List<string> files, out bool csFlag);
                var bcAlgorithm = new BinCutScriptMain(AppendText)
                {
                    BinCutFilePath = _binCutSpec,
                    BinCutPostFilePath = _postBinCutSpec,
                    TestProgramFilePath = _testProgram,
                    TestPlanFilePath = _testPlan,
                    IdsDistributionFilePath = _idsDistribution,
                    OutPutFolder = _outputFolder,
                    DataLogFiles = files,
                    ProjectName = BinCutScriptMainHelpers.GetProjectName(_binCutSpec, _postBinCutSpec, _testPlan, _testProgram),
                    Job = job
                };

                var createSettingMain = new CreateSettingMain();
                createSettingMain.CreateSettings(Path.Combine(Directory.GetCurrentDirectory(), "Settings"), bcAlgorithm.ProjectName, false);

                Response.Report($"Current Job: {job.JobType}");
                BinCutConfig.GetProjectConfig(bcAlgorithm.ProjectName);
                if (new BinCutReadMain(bcAlgorithm, AppendText).ReadBinCutData(csFlag, LocalSpecs.ProjectIniFileName))
                {
                    bcAlgorithm.DatalogValidation();
                }
                Response.Report("Finished BinCut Check Script...");

            }
            else
            {
                Response.Report("No such BinCut data files, doesn't need to execute BinCut check script report.");
            }
            Thread.Sleep(3000);
        }

        public static void AppendText(string text, Color color)
        {
            Logger logger = LogManager.GetCurrentClassLogger();
            logger.Trace(text);
        }
    }
}
