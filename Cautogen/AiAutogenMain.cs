using System.IO;
using System.Linq;

using Automation.Static;

using Cautogen.AiAutogen;
using Cautogen.AutoCZ.CharPostProcessor.Utility;

using DebugPlanReaderLib.DebugPlan;

using LogLib.Static;
using LogLib.Utility;

namespace Cautogen
{
    public class AiAutogenMain
    {
        private string job = "";
        private string project = "";
        private string channel = "";
        private string projectPatternFolder = "";
        private string charPlan = "";
        private string outputDirectory = "";
        private string baseTestProgram = "";
        private string testProgramName = "";
        private string csLibraryPath = "";
        private string resultIgxlFile = "";

        private DebugPlanMain debugAiTestPlan = null;

        public AiAutogenMain()
        {
        }

        public void Execute()
        {
            job = LocalSpecs.AllJobs.Count == 0 ? LocalSpecs.CurrentJob : LocalSpecs.AllJobs.FirstOrDefault();
            project = LocalSpecs.CurrentProject;
            channel = LocalSpecs.DefaultChannelMap;
            projectPatternFolder = LocalSpecs.PatternFolder;
            charPlan = LocalSpecs.CharPlanFileName;
            outputDirectory = LocalSpecs.TarFolder;
            baseTestProgram = string.IsNullOrEmpty(LocalSpecs.BaseTestProgram) ? "" : LocalSpecs.BaseTestProgram;
            testProgramName = LocalSpecs.TestProgramName;
            csLibraryPath = LocalSpecs.CsLibraryFolder;

            Response.Report($"Argument CurrentProjectName {project}");
            Response.Report($"Argument JobName {job}");
            Response.Report($"Argument ChannelMap {channel}");
            Response.Report($"Argument CharPlan {charPlan}");
            Response.Report($"Argument BaseTestProgram {baseTestProgram}");
            Response.Report($"Argument CSLibraryPath {csLibraryPath}");
            Response.Report($"Argument PatternFolder {projectPatternFolder}");
            Response.Report($"Argument OutputDirectory {outputDirectory}");
            Response.Report($"Argument Output TestProgramName {testProgramName}");

            string projectTimeSetFolder = Path.Combine(projectPatternFolder, "TimeSet");

            Response.Report($"Reading Pattern Info and AutoAI TestPlan...");

            outputDirectory = Path.Combine(outputDirectory, "IGLink");
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            debugAiTestPlan = new DebugPlanMain(charPlan, projectPatternFolder, baseTestProgram, true, outputDirectory);
            debugAiTestPlan.Read();
            debugAiTestPlan.CheckAll(projectPatternFolder, projectTimeSetFolder);
            Response.Report($"Finished read Pattern Info and AutoAI TestPlan...");

            resultIgxlFile = AutoProgramAiMain.Main(job, baseTestProgram, projectPatternFolder, outputDirectory, testProgramName, "", "", "", debugAiTestPlan, true, csLibraryPath);

            LogHelper.Info($"Copy csharp library to output path");
            string resultLibraryFolderPath = Path.Combine(outputDirectory, "central_library_cs", "bin");

            if (!Directory.Exists(resultLibraryFolderPath))
            {
                Directory.CreateDirectory(resultLibraryFolderPath);
            }

            FolderOperation.CopyFilesRecursively(csLibraryPath, resultLibraryFolderPath);

            LogHelper.Info($"====Result Summary====");
            LogHelper.Info($"Result IGXL: {resultIgxlFile}");
            LogHelper.Info($"Result IGXL Library: {Path.GetDirectoryName(resultLibraryFolderPath)}");
            LogHelper.Info($"Result Site0IndexFile: {debugAiTestPlan.Site0IndxFile}");
            LogHelper.Info($"Result EnableWord: {string.Join(",", debugAiTestPlan.AiTestPlanEnableWords)}");
            LogHelper.Info($"Result TesterID: {debugAiTestPlan.TesterID}");
            LogHelper.Info($"========DONE==========");
        }
    }
}
