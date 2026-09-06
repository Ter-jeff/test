using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

using AutogenCommandLine.CommandLineOptions;

using Automation;
using Automation.GenerateIgxl.PostAction.GenIgxlProj;
using Automation.GenerateIgxl.PostAction.GenT0TXProgram;
using Automation.GenerateIgxl.PostAction.GenVreTestCase;
using Automation.GenerateIgxl.PostAction.VersionTracing;
using Automation.Reader;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility;
using Automation.Utility.Basic;
using Automation.Utility.TpUpdate.HardIPBinoutTPUpdate;

using Cautogen;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.Extension;
using CommonLib.Static;
using CommonLib.Utility;

using IgxlLib.IgxlSheets;

using LcdLib;

using LogLib.Static;

using MyCommandLineLib;

using OfficeOpenXml;

using ProjectConfigLib.ProjectConfig;

using RfLib;


using TestPlanLib.PatternListCsvFile;
using TestPlanLib.Static;

namespace AutogenCommandLine.Tools.Autogen
{
    public partial class AutogenCommandLineApp : CommandLineApplicationBase
    {
        public AutogenCommandLineApp()
        {
            ToolName = "Autogen";
        }

        public override ICommandLineOptions ValidateInput(ICommandLineOptions options)
        {
            var autogenOptions = (AutogenOptions)options;
            if (!string.IsNullOrEmpty(autogenOptions.RepoFolderPath))
            {
                autogenOptions.Initialize();
            }

            UseStateMachine &= CheckExcelFilePath(autogenOptions.TestPlanFile, "-t", false);

            UseStateMachine &= CheckIniFilePath(autogenOptions.ProjectConfig, "-h", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.BasicConfigFile, "--basciconfigfile", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.BinCutInstanceNamingRule, "--binCutinstanceNamingrule", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.RtosCategoryConfig, "--rtoscategoryconfig", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.TnAssignment, "--tnassignment", false);

            UseStateMachine &= CheckXmlFilePath(autogenOptions.ScanConfig, "--scanconfig", false);

            UseStateMachine &= CheckXmlFilePath(autogenOptions.MbistConfig, "--mbistconfig", false);

            UseStateMachine &= CheckXmlFilePath(autogenOptions.SpiConfig, "--spiconfig", false);

            UseStateMachine &= CheckXmlFilePath(autogenOptions.HardipConfig, "--hardipconfig", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.BinNumberConfig, "--binnumberconfig", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.ScghFile, "-s", false);

            UseStateMachine &= CheckCsvFilePath(autogenOptions.PatternListCsvFile, "-c", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.BinCutFile, "-b", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.PostBinCutFile, "-n", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.BincutModeSequence, "-k", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.CharPlan, "--charplan", false);

            UseStateMachine &= CheckCsvFilePath(autogenOptions.VoltageTable, "-v", false);

            UseStateMachine &= CheckFolderPathOrTxt(autogenOptions.PatternInfoFile, "-g", false);

            UseStateMachine &= CheckOtpFilesPath(autogenOptions.OtpFiles, "-f", false);

            UseStateMachine &= CheckYamlFilePath(autogenOptions.YamlFile, false);

            try
            {
                UseStateMachine &= CheckFolderPath(autogenOptions.PatternFolder, "-a", false);
            }
            catch (Exception e)
            {
                if (autogenOptions.Mock != 1)
                {
                    throw new CommandLineException(ToolName, e.Message);
                }
            }

            try
            {
                UseStateMachine &= CheckFolderPath(autogenOptions.TimesetFolder, "-i", false);
            }
            catch (Exception e)
            {
                if (autogenOptions.Mock != 1)
                {
                    throw new CommandLineException(ToolName, e.Message);
                }
            }

            UseStateMachine &= CheckFolderPath(autogenOptions.BasLibraryFolder, "-d", false);

            UseStateMachine &= CheckOutputFilePath(autogenOptions.OutputDirectory);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.EqnVoltage, "-eqn", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.PowerBinning, "--powerbinning", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.EfuseTestPlan, "--efusetestplan", false);

            UseStateMachine &= CheckExcelFilePath(autogenOptions.TtrTable, "--ttr", false);

            UseStateMachine &= CheckFolderPath(autogenOptions.CustomPath, "--custom", false);

            return options;
        }

        public override ICommandLineApplication Execute(ICommandLineOptions options)
        {
            var stopWatch = new Stopwatch();
            stopWatch.Start();
            try
            {
                var args = (AutogenOptions)options;
                PreWorkFlow(args);

                if (LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    GenerateIgxlMainRf.Run();
                }
                else if (LocalSpecs.Options.Device == EnumDevice.LCD)
                {
                    new GenerateIgxlMainLcd().Run();
                }
                else
                {
                    new GenerateIgxlMain().Run();
                }

                PostWorkFlow(args.Mock, args.OutputDirectory);
                LocalSpecs.ElapsedTimeResults.WriteToJson();
            }
            catch (Exception e)
            {
                Response.Report(e.StackTrace);
                throw new Exception(e.StackTrace);
            }
            finally
            {
                stopWatch.Stop();
                Response.Report(string.Format("Total Process Time : " + TimeSpan.FromMilliseconds(stopWatch.ElapsedMilliseconds).ToString(@"hh\:mm\:ss"), EnumMessageLevel.General));
            }
            return this;
        }

        private static void GenerateIgxl(List<FileInfo> fileInfos)
        {
            IIgxlPackager packager = IgxlPackagerFactory.Create(skipIgLink: LocalSpecs.Options.SkipIgLink);

            Response.Report("Generating IGXL Test Program ...", EnumMessageLevel.CheckPoint);
            var files = fileInfos.Select(x => x.FullName).ToList();
            files.AddRange(TestProgram.IgxlWorkBk.AllIgxlSheets.Keys.Select(x => x + ".txt"));
            files.AddRange(TestProgram.NonIgxlSheetsList.SheetList.Select(x => x + ".txt"));
            var fileList = fileInfos.Select(x => x.FullName).ToList();
            object[] bypassedSheets1 = [typeof(MainFlow), typeof(JobListSheet)];
            var fileNames1 = TestProgram.IgxlWorkBk.AllIgxlSheets
              .Where(kvp => !bypassedSheets1.Any(type => kvp.Value.GetType() == (Type)type))
              .Select(kvp => kvp.Key + ".txt")
              .ToList();
            fileList.AddRange(fileNames1);
            fileList.AddRange(TestProgram.NonIgxlSheetsList.SheetList.Select(x => x + ".txt"));
            fileList.AddRange(TestProgram.SubProgIgxlWorkBk.AllIgxlSheets.Keys.Select(x => x + ".txt"));
            var t0TxFiles = fileInfos.Select(x => x.FullName).ToList();
            object[] bypassedSheets = [typeof(InstanceSheet), typeof(SubFlowSheet), typeof(MainFlow), typeof(JobListSheet)];
            var fileNames = TestProgram.IgxlWorkBk.AllIgxlSheets
                .Where(kvp => !bypassedSheets.Any(type => kvp.Value.GetType() == (Type)type))
                .Select(kvp => kvp.Key + ".txt")
                .ToList();
            t0TxFiles.AddRange(fileNames);
            t0TxFiles.AddRange(TestProgram.NonIgxlSheetsList.SheetList.Select(x => x + ".txt"));
            t0TxFiles.AddRange(TestProgram.T0TxIgxlWorkBk.AllIgxlSheets.Keys.Select(x => x + ".txt"));
            var oasisMergeSheets = files.Where(File.Exists).ToList();
            var subOasisMergeSheets = fileList.Where(File.Exists).ToList();
            var t0TxOasisMergeSheets = t0TxFiles.Where(File.Exists).ToList();

            // Inject "#Const isUFP = True/False" at the head of every .bas before the packager runs.
            // Mutation is idempotent and must happen even when the packager is a no-op.
            UfpBasMutator.Apply(oasisMergeSheets.Concat(subOasisMergeSheets).Concat(t0TxOasisMergeSheets).Distinct());

            packager.GenIgxlProg(
                oasisMergeSheets,
                LocalSpecs.TarFolder,
                LocalSpecs.TestProgramName,
                TestProgram.IgxlWorkBk,
                LocalSpecs.IsUnitTest
            );

            if (TestProgram.SubProgIgxlWorkBk.MainFlowSheets.Count != 0)
            {
                packager.GenIgxlProg(
                    subOasisMergeSheets,
                    LocalSpecs.TarFolder,
                    LocalSpecs.TestProgramName + "_Sub",
                    TestProgram.IgxlWorkBk,
                    LocalSpecs.IsUnitTest
                );
            }
            if (LocalSpecs.Options.GenerateT0TXTestprogram && TestProgram.T0TxIgxlWorkBk.MainFlowSheets.Count != 0)
            {
                packager.GenIgxlProg(
                    t0TxOasisMergeSheets,
                    LocalSpecs.TarFolder,
                    LocalSpecs.TestProgramName + "_T0Tx",
                    TestProgram.IgxlWorkBk,
                    LocalSpecs.IsUnitTest
                );
            }
            Response.Report("Automation ModuleMain Completed!",
                EnumMessageLevel.CheckPoint);
            Response.Report(
                Environment.NewLine
                    + "===================================================="
            );
        }

        private static bool IsFileNameApplicable(string str)
        {
            return !str.Equals("N/A") && !string.IsNullOrEmpty(str);
        }

        private static string GetValidFileName(string str)
        {
            return str.Equals(@"N\A", StringComparison.CurrentCultureIgnoreCase) ? "" : str;
        }

        private static void GenerateErrorReport(int mock)
        {
            if (mock == 1)
            {
                return;
            }
            if (IsFileNameApplicable(LocalSpecs.TestPlanFileName))
            {
                string originalTestPlan = Path.GetFullPath(LocalSpecs.TestPlanFileName);
                string originalTestPlanName = Path.GetFileName(LocalSpecs.TestPlanFileName);
                var fileInfo = new FileInfo(originalTestPlanName);
                string copiedTestPlanName = originalTestPlanName.Replace(
                    fileInfo.Extension,
                    "_Real" + VersionControl.Timestamp + fileInfo.Extension
                );
                string outPlanName = Path.Combine(FolderStructure.DirIgLink, copiedTestPlanName);
                outPlanName = FileManager.CopyFile(originalTestPlan, outPlanName);
                LocalSpecs.TestPlanFileName = outPlanName;
            }
            if (IsFileNameApplicable(LocalSpecs.ScghFileName))
            {
                string originalScgh = Path.GetFullPath(LocalSpecs.ScghFileName);
                string originalScghName = Path.GetFileName(LocalSpecs.ScghFileName);
                var fileInfo = new FileInfo(originalScghName);
                string copiedScghName = originalScghName.Replace(
                    fileInfo.Extension,
                    "_Real" + VersionControl.Timestamp + fileInfo.Extension
                );
                string outScghName = Path.Combine(FolderStructure.DirIgLink, copiedScghName);
                outScghName = FileManager.CopyFile(originalScgh, outScghName);
                LocalSpecs.ScghFileName = outScghName;
            }
            if (IsFileNameApplicable(LocalSpecs.BinCutFileName))
            {
                string originalBinCutFile = Path.GetFullPath(LocalSpecs.BinCutFileName);
                string originalBinCutFileName = Path.GetFileName(LocalSpecs.BinCutFileName);
                var fileInfo = new FileInfo(originalBinCutFileName);
                string copiedBinCutFileName = originalBinCutFileName.Replace(
                    fileInfo.Extension,
                    "_Real" + VersionControl.Timestamp + fileInfo.Extension
                );
                string outBinCutFileName = Path.Combine(FolderStructure.DirIgLink, copiedBinCutFileName);
                outBinCutFileName = FileManager.CopyFile(originalBinCutFile, outBinCutFileName);
                LocalSpecs.BinCutFileName = outBinCutFileName;
            }
            if (IsFileNameApplicable(LocalSpecs.BinCutPostFileName))
            {
                string originalPbcFile = Path.GetFullPath(LocalSpecs.BinCutPostFileName);
                string originalPbcFileName = Path.GetFileName(LocalSpecs.BinCutPostFileName);
                var fileInfo = new FileInfo(originalPbcFileName);
                string copiedPbcFileName = originalPbcFileName.Replace(
                    fileInfo.Extension,
                    "_Real" + VersionControl.Timestamp + fileInfo.Extension
                );
                string outPbcFileName = Path.Combine(FolderStructure.DirIgLink, copiedPbcFileName);
                outPbcFileName = FileManager.CopyFile(originalPbcFile, outPbcFileName);
                LocalSpecs.BinCutPostFileName = outPbcFileName;
            }
            if (IsFileNameApplicable(LocalSpecs.BinCutModeSeqFileName))
            {
                string originalBmsFile = Path.GetFullPath(LocalSpecs.BinCutModeSeqFileName);
                string originalBmsFileName = Path.GetFileName(LocalSpecs.BinCutModeSeqFileName);
                var fileInfo = new FileInfo(originalBmsFileName);
                string copiedBmsFileName = originalBmsFileName.Replace(
                    fileInfo.Extension,
                    "_Real" + VersionControl.Timestamp + fileInfo.Extension
                );
                string outBmsFileName = Path.Combine(FolderStructure.DirIgLink, copiedBmsFileName);
                outBmsFileName = FileManager.CopyFile(originalBmsFile, outBmsFileName);
                LocalSpecs.BinCutModeSeqFileName = outBmsFileName;
            }

            string planCopyFile = GetValidFileName(LocalSpecs.TestPlanFileName);
            string scghCopyFile = GetValidFileName(LocalSpecs.ScghFileName);
            string binCutCopyFile = GetValidFileName(LocalSpecs.BinCutFileName);
            string postBinCutCopyFile = GetValidFileName(LocalSpecs.BinCutPostFileName);
            string modeSequenceCopyFile = GetValidFileName(LocalSpecs.BinCutModeSeqFileName);
            var copyFiles = new List<string>
            {
                planCopyFile,
                scghCopyFile,
                binCutCopyFile,
                postBinCutCopyFile,
                modeSequenceCopyFile
            }.Where(f => !string.IsNullOrEmpty(f) && File.Exists(f)).ToList();

            if (copyFiles.Count != 0)
            {
                using var pkg = new ExcelPackage(new FileInfo(copyFiles[0]));
                ErrorReportManager.GenErrorReport(pkg, [.. copyFiles.Skip(1)], "ErrorReport");
                pkg.Compression = CompressionLevel.BestSpeed;
                pkg.Save();
            }
        }

        protected static void PostWorkFlow(int mock, string outputFolder)
        {
            if (LocalSpecs.BinOutReportFileName != "N/A")
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection("Binout", () =>
                {
                    Response.Report("Start hardip bin out tp update ...");
                    new HardIpBinOutTpUpdateMain(LocalSpecs.BinOutReportFileName, TestProgram.IgxlWorkBk.SubFlowSheets, [.. TestProgram.IgxlWorkBk.BinTblSheets.Values]).WorkFlow();
                    Response.Report("Hardip bin out tp update done.");
                });
            }

            //Generate T0TX test program
            if (LocalSpecs.Options.GenerateT0TXTestprogram)
            {
                if (TestProgram.T0TxIgxlWorkBk.MainFlowSheets.Count != 0)
                {
                    LocalSpecs.ElapsedTimeResults.MeasureSection("T0tx", () =>
                    {
                        var genT0TXProgram = new GenT0TXProgramMain(TestProgram.IgxlWorkBk.SubFlowSheets, TestProgram.IgxlWorkBk.InsSheets);
                        genT0TXProgram.Workflow();
                        TestProgram.T0TxIgxlWorkBk.PrintAllSheets();
                    });
                }
            }

            if (!string.IsNullOrEmpty(LocalSpecs.CharPlanFileName) && !LocalSpecs.CharPlanFileName.Equals(@"N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                LocalSpecs.ElapsedTimeResults.MeasureSection("Cautogen", () =>
                {
                    Response.Report("Start C-Autogen ...");
                    new CharAutogenMain().Execute();
                    Response.Report(" C-Autogen done.");
                });
            }
            #region VreTestCaseTable
            if (LocalSpecs.Options.VreEnable && TestPlanStatic.VreTestCaseTable != null)
            {
                Response.Report("Generating VreTestCaseTable ~", EnumMessageLevel.CheckPoint);
                new GenVreTestCaseMain(TestPlanStatic.VreTestCaseTable,
                                        TestPlanStatic.MappingCoreTable,
                                        TestPlanStatic.HarvestingTruthTableSheets.FirstOrDefault(),
                                        TestPlanStatic.VreMbistLookupTable,
                                        TestPlanStatic.FlagOperationSheets).Workflow();
                Response.Report("VreTestCaseTable Completed !", EnumMessageLevel.EndPoint);
            }
            #endregion
            if (!LocalSpecs.CheckOnly)
            {
                VersionTracingMain.WorkFlow();

                GenerateIgxl(GetVbFiles());
            }
            List<Error> errors = ErrorReportManager.GetSortedErrors();
            int errorCounts = errors.Count(x => x.ErrorLevel == EnumErrorLevel.Error);
            int warningCounts = errors.Count(x => x.ErrorLevel == EnumErrorLevel.Warning);
            if (LocalSpecs.CheckOnly && Cautogen.AutoCZ.CharPreProcessor.ReportManager.ErrorManager.ErrorListDict != null)
            {
                List<Cautogen.AutoCZ.CharPreProcessor.ReportManager.ErrorMessage> charPlanErrors =
                    Cautogen.AutoCZ.CharPreProcessor.ReportManager.ErrorManager.ErrorListDict.Values.SelectMany(x => x).ToList();
                int charPlanErrorCounts = charPlanErrors.Count(x => x.ErrorLevel == Cautogen.AutoCZ.CharPreProcessor.ReportManager.ErrorLevel.Error);
                int charPlanWarningCounts = charPlanErrors.Count(x => x.ErrorLevel == Cautogen.AutoCZ.CharPreProcessor.ReportManager.ErrorLevel.Warning);
                Response.Report($"Char plan check summary: {charPlanErrorCounts} Error(s); {charPlanWarningCounts} Warning(s); see the char plan error report for details.", EnumMessageLevel.CheckPoint);
            }
            else
            {
                Response.Report($"The summary result is: {errorCounts} Error(s); {warningCounts} Warning(s);", EnumMessageLevel.CheckPoint);
            }

            GenerateErrorReport(mock);

            if (mock != 1 && !string.IsNullOrEmpty(outputFolder))
            {
                string outputFile = Path.Combine(outputFolder, "Validaton_ErrorReport.xlsx");
                using var package = new ExcelPackage(new FileInfo(outputFile));
                if (errors.Count != 0)
                {
                    ErrorReportManager.GenErrorReport(package, "ErrorReport");
                }
                if (package.Workbook.Worksheets.Count > 0)
                {
                    package.Compression = CompressionLevel.BestSpeed;
                    package.Save();
                }
            }
            else
            {
                string report = Path.Combine(outputFolder, "Error.txt");
                if (ErrorReportManager.GetErrorList().Count != 0)
                {
                    IEnumerable<string> lines = errors.Take(20000).Select(x => x.Print());
                    File.AppendAllLines(report, lines);
                }
            }

            Response.Report("T/P output path:  " + FolderStructure.DirIgLink, EnumMessageLevel.Result);
            if (!LocalSpecs.CheckOnly)
            {
                Response.Report("Generating IGXL Completed!", EnumMessageLevel.EndPoint);
            }
            else
            {
                Response.Report("Skipped IGXL generation (Check-Only mode).", EnumMessageLevel.EndPoint);
            }
        }

        [GeneratedRegex("(_EfusePlan|_EFUSE_EXTERNAL).*", RegexOptions.IgnoreCase)]
        private static partial Regex EfusePlanRegex();

        [GeneratedRegex("_Test.*Plan.*", RegexOptions.IgnoreCase)]
        private static partial Regex TestPlanRegex();

        [GeneratedRegex("_scgh", RegexOptions.IgnoreCase)]
        private static partial Regex ScghRegex();

        [GeneratedRegex("_Pattern.*_.*.csv", RegexOptions.IgnoreCase)]
        private static partial Regex PatternListCsvRegex();

        [GeneratedRegex("_TW_V", RegexOptions.IgnoreCase)]
        private static partial Regex TwVRegex();

        [GeneratedRegex(@"((\w+)_)?(VoltageTable|Testsetting|VoTa|VolTa)(_\w+)?_(CP|WLFT|FT|FQA|RMA|EMA|T0TxFT)[\d]+(_\w+)?", RegexOptions.IgnoreCase)]
        private static partial Regex VoltageTableRegex();

        [GeneratedRegex("(Bin_Cut|Voltage_Binning)", RegexOptions.IgnoreCase)]
        private static partial Regex BinCutRegex();

        [GeneratedRegex("(Post_BinCut|PBC)", RegexOptions.IgnoreCase)]
        private static partial Regex PostBinCutRegex();

        [GeneratedRegex("(bincut_mode_sequence|BMS)", RegexOptions.IgnoreCase)]
        private static partial Regex BinCutModeSequenceRegex();

        [GeneratedRegex("(EquationBasedVoltages|EQN)", RegexOptions.IgnoreCase)]
        private static partial Regex EquationBasedVoltagesRegex();

        [GeneratedRegex("_TTR_", RegexOptions.IgnoreCase)]
        private static partial Regex TtrRegex();

        [GeneratedRegex("(PowerBinning|PowerScreening)", RegexOptions.IgnoreCase)]
        private static partial Regex PowerBinningRegex();

        [GeneratedRegex("DRAM", RegexOptions.IgnoreCase)]
        private static partial Regex DramRegex();

        [GeneratedRegex("FuseCheck", RegexOptions.IgnoreCase)]
        private static partial Regex FuseCheckRegex();

        private static Dictionary<string, string> GetFilesFromDirectory(string dirPath)
        {
            var resultDic = new Dictionary<string, string>();
            string[] fileList = Directory.GetFiles(dirPath);
            var voltageTableList = new List<string>();
            var ttrTableList = new List<string>();
            var binCutFileList = new List<string>();
            var powerBinningList = new List<string>();
            foreach (string file in fileList)
            {
                string fileName = Path.GetFileName(file);
                if (fileName.StartsWith('~'))
                {
                    continue;
                }

                if (EfusePlanRegex().IsMatch(fileName))
                {
                    resultDic.Add("Efuse Test Plan", file);
                }
                else if (TestPlanRegex().IsMatch(fileName))
                {
                    resultDic.Add("Test Plan", file);
                }
                else if (ScghRegex().IsMatch(fileName))
                {
                    resultDic.Add("SCGH", file);
                }
                else if (PatternListCsvRegex().IsMatch(fileName) && !TwVRegex().IsMatch(fileName))
                {
                    resultDic.Add("Pattern List Csv", file);
                }
                else if (VoltageTableRegex().IsMatch(fileName))
                {
                    voltageTableList.Add(file);
                }
                else if (BinCutRegex().IsMatch(fileName))
                {
                    binCutFileList.Add(file);
                }
                else if (PostBinCutRegex().IsMatch(fileName))
                {
                    resultDic.Add("Post Bin Cut", file);
                }
                else if (BinCutModeSequenceRegex().IsMatch(fileName))
                {
                    resultDic.Add("BinCut Mode Sequence", file);
                }
                else if (EquationBasedVoltagesRegex().IsMatch(fileName))
                {
                    resultDic.Add("Equation Based Voltages", file);
                }
                else if (TtrRegex().IsMatch(fileName))
                {
                    ttrTableList.Add(file);
                }
                else if (PowerBinningRegex().IsMatch(fileName))
                {
                    powerBinningList.Add(file);
                }
                else if (DramRegex().IsMatch(fileName))
                {
                    resultDic.Add("DRAM Type Table", file);
                }
                else if (FuseCheckRegex().IsMatch(fileName))
                {
                    resultDic.Add("Fuse Check Table", file);
                }
            }
            if (binCutFileList.Count != 0)
            {
                resultDic.Add("Bin Cut", string.Join(",", binCutFileList.Distinct()));
            }

            if (voltageTableList.Count != 0)
            {
                resultDic.Add("Voltage Tables", string.Join(",", voltageTableList.Distinct()));
            }

            if (ttrTableList.Count != 0)
            {
                resultDic.Add("HIP TTR Table", string.Join(",", ttrTableList.Distinct()));
            }

            if (powerBinningList.Count != 0)
            {
                resultDic.Add("PowerBinning", string.Join(",", powerBinningList.Distinct()));
            }

            return resultDic;
        }

        public static (List<string> tpSheetsList, List<string> scghSheetsList) PreWorkFlow(AutogenOptions autogenOptions)
        {
            ResolveConfigPaths(autogenOptions);

            var inputDic = new Dictionary<string, string>();
            var selectedFiles = new List<string>();

            if (autogenOptions.Mock == 1)
            {
                Mock();
            }

            AllStatic.Clear();
            LocalSpecs.IsUnitTest = autogenOptions.Mock == 1;

            if (!string.IsNullOrEmpty(autogenOptions.InputDirectory))
            {
                inputDic = GetFilesFromDirectory(autogenOptions.InputDirectory);
            }

            SetByInputFolder(inputDic);

            InitProjectConfigAndOptions(autogenOptions);

            LocalSpecs.TestProgramName = string.IsNullOrEmpty(autogenOptions.TestProgramName) ? autogenOptions.CurrentProjectName : autogenOptions.TestProgramName;
            LocalSpecs.CurrentProject = autogenOptions.CurrentProjectName;
            LocalSpecs.TarFolder = autogenOptions.OutputDirectory;

            LoadTestPlanWorkbook(autogenOptions, selectedFiles);

            SetPatternOptions(autogenOptions);

            LocalSpecs.PatternListCsvFileName = string.IsNullOrEmpty(autogenOptions.PatternListCsvFile) ? LocalSpecs.PatternListCsvFileName : autogenOptions.PatternListCsvFile;
            if (!string.IsNullOrEmpty(LocalSpecs.PatternListCsvFileName) && !LocalSpecs.PatternListCsvFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.PatternListCsvFileName);
            }

            AllSingleton.Initialize();

            CompilePatternListCsv();

            CollectReportAndPlanFiles(autogenOptions, selectedFiles);

            LocalSpecs.BinCutFileName = string.IsNullOrEmpty(autogenOptions.BinCutFile) ? LocalSpecs.BinCutFileName : autogenOptions.BinCutFile;

            CollectBinCutRelatedFiles(autogenOptions, selectedFiles);

            CollectTtrAndPowerBinningFiles(autogenOptions, selectedFiles);

            if (!string.IsNullOrEmpty(LocalSpecs.EfuseTestPlanFileName) && !LocalSpecs.EfuseTestPlanFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.EfuseTestPlanFileName);
            }

            SetLibraryFolders(autogenOptions);

            SetMiscFileOptions(autogenOptions);

            SetSearchValidPatternRev(autogenOptions);

            SetSettingAndConfigFiles(autogenOptions);

            if (!string.IsNullOrEmpty(LocalSpecs.ScghFileName))
            {
                var package = new ExcelPackage(new FileInfo(LocalSpecs.ScghFileName));
                EpWorkbook.ScghWorkbook = package.Workbook;
            }

            ProcessBinCutFiles(autogenOptions, selectedFiles);

            LoadBinCutWorkbooks();

            ProcessVoltageTables(autogenOptions, selectedFiles);

            if (!string.IsNullOrEmpty(autogenOptions.BasicConfigFile))
            {
                var package = new ExcelPackage(new FileInfo(autogenOptions.BasicConfigFile));
                SettingStatic.BasicConfigWorkbook = package.Workbook;
            }

            var assembly = Assembly.GetExecutingAssembly();
            string version = AssemblyProvider.Current.GetFileVersion(FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion);
            LocalSpecs.AutogenVer = version;
            LocalSpecs.SelectedFiles.AddRange(selectedFiles);
            List<string> tpSheetsList = EpWorkbook.TestPlanWorkbook.GetMatchPlanSheets(".*");
            List<string> scghSheetsList = EpWorkbook.ScghWorkbook == null ? [] : EpWorkbook.ScghWorkbook.GetMatchPlanSheets(".*");
            List<string> binCutList = EpWorkbook.BinCutWorkbook == null ? [] : EpWorkbook.BinCutWorkbook.GetMatchPlanSheets(".*");
            List<string> postBinCutList = EpWorkbook.BinCutPostWorkbook == null ? [] : EpWorkbook.BinCutPostWorkbook.GetMatchPlanSheets(".*");
            BlockStatus.SetAutomation(tpSheetsList, scghSheetsList, LocalSpecs.PatternListCsvFileName, binCutList, postBinCutList, autogenOptions.OtpFiles, autogenOptions.YamlFile);
            FolderStructure.CreateFolder();
            return (tpSheetsList, scghSheetsList);
        }

        private static void ResolveConfigPaths(AutogenOptions autogenOptions)
        {
            string repoProjectRepoPath = autogenOptions.RepoFolderPath;
            string projectName = autogenOptions.CurrentProjectName;

            autogenOptions.BasicConfigFile = UserInfo.ResolveConfigPath(autogenOptions.BasicConfigFile, repoProjectRepoPath, projectName, Path.Combine("Settings", "Basic", "Basic_Configure_{0}.xlsx"), Path.Combine("Settings", "Basic", "Basic_Configure_Default.xlsx"));
            autogenOptions.BinCutInstanceNamingRule = UserInfo.ResolveConfigPath(autogenOptions.BinCutInstanceNamingRule, repoProjectRepoPath, projectName, Path.Combine("Settings", "SCGH", "BinCutInstanceNamingRule_{0}.xlsx"), Path.Combine("Settings", "SCGH", "BinCutInstanceNamingRule_Default.xlsx"));
            autogenOptions.RtosCategoryConfig = UserInfo.ResolveConfigPath(autogenOptions.RtosCategoryConfig, repoProjectRepoPath, projectName, Path.Combine("Settings", "Basic", "RtosCategory_{0}.xlsx"), Path.Combine("Settings", "Basic", "RtosCategory_Default.xlsx"));
            autogenOptions.ScanConfig = UserInfo.ResolveConfigPath(autogenOptions.ScanConfig, repoProjectRepoPath, projectName, Path.Combine("Settings", "SCGH", "Scan_Config_{0}.xml"), Path.Combine("Settings", "SCGH", "Scan_Config_Default.xml"));
            autogenOptions.MbistConfig = UserInfo.ResolveConfigPath(autogenOptions.MbistConfig, repoProjectRepoPath, projectName, Path.Combine("Settings", "SCGH", "Mbist_Config_{0}.xml"), Path.Combine("Settings", "SCGH", "Mbist_Config_Default.xml"));
            autogenOptions.SpiConfig = UserInfo.ResolveConfigPath(autogenOptions.SpiConfig, repoProjectRepoPath, projectName, Path.Combine("Settings", "SCGH", "Spi_Config_{0}.xml"), Path.Combine("Settings", "SCGH", "Spi_Config_Default.xml"));
            autogenOptions.HardipConfig = UserInfo.ResolveConfigPath(autogenOptions.HardipConfig, repoProjectRepoPath, projectName, Path.Combine("Settings", "SCGH", "HardIP_Config_{0}.xml"), Path.Combine("Settings", "SCGH", "HardIP_Config_Default.xml"));
            autogenOptions.TnAssignment = UserInfo.ResolveConfigPath(autogenOptions.TnAssignment, repoProjectRepoPath, projectName, Path.Combine("Settings", "Basic", "TN_Assignment_{0}.xlsx"), Path.Combine("Settings", "Basic", "TN_Assignment_Default.xlsx"));
            autogenOptions.BinNumberConfig = UserInfo.ResolveConfigPath(autogenOptions.BinNumberConfig, repoProjectRepoPath, projectName, Path.Combine("Config", "BinNumberConfig_{0}.xlsx"), Path.Combine("Config", "BinNumberConfig.xlsx"));
            autogenOptions.ProjectConfig = UserInfo.ResolveConfigPath(autogenOptions.ProjectConfig, repoProjectRepoPath, projectName, Path.Combine("Settings", "ProjectConfig_{0}.ini"), Path.Combine("Settings", "ProjectConfig_OTC.ini"));
            autogenOptions.NeedSheet = UserInfo.ResolveConfigPath(autogenOptions.NeedSheet, repoProjectRepoPath, projectName, Path.Combine("Settings", "Basic", "NeedSheetConfig_{0}.xml"), Path.Combine("Settings", "Basic", "NeedSheetConfig_Default.xml"));
        }

        private static void InitProjectConfigAndOptions(AutogenOptions autogenOptions)
        {
            LocalSpecs.ProjectIniFileName = string.IsNullOrEmpty(autogenOptions.ProjectConfig) ? "" : autogenOptions.ProjectConfig;
            if (!string.IsNullOrEmpty(LocalSpecs.ProjectIniFileName))
            {
                ProjectConfigSingleton.Instance().InitializeProjectConfigSetting();
                ProjectConfigSingleton.Instance().LoadProjectConfig(LocalSpecs.ProjectIniFileName);
            }
            var optionals = new Options(ProjectConfigSingleton.Instance());
            LocalSpecs.Options = optionals;

            NeededSheets.InitSheetName(LocalSpecs.Options.Device, autogenOptions.RepoFolderPath, autogenOptions.CurrentProjectName, autogenOptions.NeedSheet);
        }

        private static void LoadTestPlanWorkbook(AutogenOptions autogenOptions, List<string> selectedFiles)
        {
            LocalSpecs.TestPlanFileName = string.IsNullOrEmpty(autogenOptions.TestPlanFile) ? LocalSpecs.TestPlanFileName : autogenOptions.TestPlanFile;
            if (!string.IsNullOrEmpty(LocalSpecs.TestPlanFileName) && !LocalSpecs.TestPlanFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.TestPlanFileName);
            }
            LocalSpecs.EfuseTestPlanFileName = string.IsNullOrEmpty(autogenOptions.EfuseTestPlan) ? LocalSpecs.EfuseTestPlanFileName : autogenOptions.EfuseTestPlan;
            if (!string.IsNullOrEmpty(LocalSpecs.TestPlanFileName))
            {
                var package = new ExcelPackage(new FileInfo(LocalSpecs.TestPlanFileName));
                EpWorkbook.TestPlanWorkbook = package.Workbook;
                if (!string.IsNullOrEmpty(LocalSpecs.EfuseTestPlanFileName) && LocalSpecs.EfuseTestPlanFileName != "N/A")
                {
                    ExcelWorkbook efuseTestPlan = new ExcelPackage(new FileInfo(LocalSpecs.EfuseTestPlanFileName)).Workbook;
                    foreach (ExcelWorksheet sheet in efuseTestPlan.Worksheets)
                    {
                        if (EpWorkbook.TestPlanWorkbook.Worksheets.Any(x => x.Name.Equals(sheet.Name)))
                        {
                            EpWorkbook.TestPlanWorkbook.Worksheets.Delete(sheet.Name);
                        }
                        EpWorkbook.TestPlanWorkbook.Worksheets.Add(sheet.Name, sheet);
                    }
                }
            }
        }

        private static void SetPatternOptions(AutogenOptions autogenOptions)
        {
            LocalSpecs.TimeSetFolder = string.IsNullOrEmpty(autogenOptions.TimesetFolder) ? "" : autogenOptions.TimesetFolder;
            LocalSpecs.PatternFolder = string.IsNullOrEmpty(autogenOptions.PatternFolder) ? "" : autogenOptions.PatternFolder;
            LocalSpecs.Options.IsIgnorePatternCheck = !string.IsNullOrEmpty(autogenOptions.SkipPatCheck) && autogenOptions.SkipPatCheck.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
            LocalSpecs.Options.SkipIgLink = !string.IsNullOrEmpty(autogenOptions.SkipIgLink) && autogenOptions.SkipIgLink.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
            LocalSpecs.Options.GenerateT0TXTestprogram = !string.IsNullOrEmpty(autogenOptions.GenT0TX) && autogenOptions.GenT0TX.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
            LocalSpecs.Options.VreEnable = !string.IsNullOrEmpty(autogenOptions.VreEnable) && autogenOptions.VreEnable.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
            LocalSpecs.CompilePatFileName = string.IsNullOrEmpty(autogenOptions.CompilePat) ? "" : autogenOptions.CompilePat;
        }

        private static void CompilePatternListCsv()
        {
            #region PatternList CSV Compile
            bool isUseTheLatestPatVersion = !LocalSpecs.Options.IsIgnorePatternCheck && LocalSpecs.SearchValidPatternRev;
            if (LocalSpecs.PatternListCsvFileName != "N/A")
            {
                var fileInfo = new FileInfo(LocalSpecs.PatternListCsvFileName);
                var patListCsv = new InputPatternListCsv(fileInfo);
                patListCsv.Compile(LocalSpecs.CompilePatFileName, LocalSpecs.PatternFolder, LocalSpecs.TimeSetFolder, isUseTheLatestPatVersion, out Dictionary<string, CompileItem> dicCompileItem, null);
                if (dicCompileItem != null)
                {
                    LocalSpecs.CompileItem = dicCompileItem;
                }

                LocalSpecs.PatternListCsvFileName = patListCsv.FullName;
            }
            #endregion
        }

        private static void CollectReportAndPlanFiles(AutogenOptions autogenOptions, List<string> selectedFiles)
        {
            LocalSpecs.BinOutReportFileName = autogenOptions.BinoutReport;
            if (!string.IsNullOrEmpty(LocalSpecs.BinOutReportFileName) && !LocalSpecs.BinOutReportFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.BinOutReportFileName);
            }
            LocalSpecs.CharPlanFileName = autogenOptions.CharPlan;
            if (!string.IsNullOrEmpty(LocalSpecs.CharPlanFileName) && !LocalSpecs.CharPlanFileName.Equals(@"N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.CharPlanFileName);
            }
            LocalSpecs.ScghFileName = string.IsNullOrEmpty(autogenOptions.ScghFile) ? LocalSpecs.ScghFileName : autogenOptions.ScghFile;
            if (!string.IsNullOrEmpty(LocalSpecs.ScghFileName) && !LocalSpecs.ScghFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.ScghFileName);
            }
        }

        private static void CollectBinCutRelatedFiles(AutogenOptions autogenOptions, List<string> selectedFiles)
        {
            LocalSpecs.BinCutPostFileName = string.IsNullOrEmpty(autogenOptions.PostBinCutFile) ? LocalSpecs.BinCutPostFileName : autogenOptions.PostBinCutFile;
            if (!string.IsNullOrEmpty(LocalSpecs.BinCutPostFileName) && !LocalSpecs.BinCutPostFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.BinCutPostFileName);
            }

            LocalSpecs.BinCutModeSeqFileName = string.IsNullOrEmpty(autogenOptions.BincutModeSequence) ? LocalSpecs.BinCutModeSeqFileName : autogenOptions.BincutModeSequence;
            if (!string.IsNullOrEmpty(LocalSpecs.BinCutModeSeqFileName) && !LocalSpecs.BinCutModeSeqFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.BinCutModeSeqFileName);
            }

            LocalSpecs.EquationVoltagesFileName = string.IsNullOrEmpty(autogenOptions.EqnVoltage) ? LocalSpecs.EquationVoltagesFileName : autogenOptions.EqnVoltage;
            if (!string.IsNullOrEmpty(LocalSpecs.EquationVoltagesFileName) && !LocalSpecs.EquationVoltagesFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.EquationVoltagesFileName);
            }
        }

        private static void CollectTtrAndPowerBinningFiles(AutogenOptions autogenOptions, List<string> selectedFiles)
        {
            LocalSpecs.TtrSummaryFileName = string.IsNullOrEmpty(autogenOptions.TtrTable) ? LocalSpecs.TtrSummaryFileName : autogenOptions.TtrTable;
            if (LocalSpecs.TtrSummaryFileName.Contains(','))
            {
                string[] ttrSumFiles = LocalSpecs.TtrSummaryFileName.Split(',');
                using var package = new ExcelPackage(new FileInfo(ttrSumFiles[0]));
                for (int i = 1; i < ttrSumFiles.Length; i++)
                {
                    ExcelWorkbook ttrFile = new ExcelPackage(new FileInfo(ttrSumFiles[i])).Workbook;
                    foreach (ExcelWorksheet sheet in ttrFile.Worksheets)
                    {
                        if (!package.Workbook.Worksheets.Any(x => x.Name.Equals(sheet.Name)))
                        {
                            package.Workbook.Worksheets.Add(sheet.Name, sheet);
                        }
                    }
                }
                string merge = ttrSumFiles[0].Replace(".xlsx", "_" + TimeContext.Now.ToString("yyMMddHHmm") + "_Merge.xlsx");
                package.SaveAs(new FileInfo(merge));
                LocalSpecs.TtrSummaryFileName = merge;
            }

            if (!string.IsNullOrEmpty(LocalSpecs.TtrSummaryFileName) && !LocalSpecs.TtrSummaryFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.TtrSummaryFileName);
            }

            if (!string.IsNullOrEmpty(autogenOptions.PowerBinning))
            {
                LocalSpecs.AllPowerBinningFileName = [];
                foreach (string filePath in autogenOptions.PowerBinning.Split(',').ToList())
                {
                    LocalSpecs.AllPowerBinningFileName.Add(filePath);
                    selectedFiles.Add(filePath);
                }
            }
        }

        private static void SetLibraryFolders(AutogenOptions autogenOptions)
        {
            LocalSpecs.CsLibraryFolder = string.IsNullOrEmpty(autogenOptions.CsLibraryPath) ? "" : autogenOptions.CsLibraryPath;
            if (!string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                if (string.IsNullOrEmpty(autogenOptions.BasLibraryFolder))
                {
                    LocalSpecs.BasLibraryFolder = Path.Combine(Directory.GetCurrentDirectory(), "Settings", "DefaultVbtForCs");
                }
                else
                {
                    LocalSpecs.BasLibraryFolder = autogenOptions.BasLibraryFolder;
                }
            }
            else
            {
                LocalSpecs.BasLibraryFolder = string.IsNullOrEmpty(autogenOptions.BasLibraryFolder) ? "" : autogenOptions.BasLibraryFolder;
            }
        }

        private static void SetMiscFileOptions(AutogenOptions autogenOptions)
        {
            LocalSpecs.HardIpInfoFileName = string.IsNullOrEmpty(autogenOptions.PatternInfoFile) ? "" : autogenOptions.PatternInfoFile;
            LocalSpecs.MbistInfoFileName = string.IsNullOrEmpty(autogenOptions.MbistInfo) ? "" : autogenOptions.MbistInfo;
            LocalSpecs.IsUfp = autogenOptions.IsUfp?.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase) ?? LocalSpecs.Options.IsUfp;

            LocalSpecs.DefaultChannelMap = string.IsNullOrEmpty(autogenOptions.DefaultChannelMap) ? "" : autogenOptions.DefaultChannelMap;
            if (!string.IsNullOrEmpty(autogenOptions.CustomPath))
            {
                LocalSpecs.CustomPath = [.. autogenOptions.CustomPath.Split(',').ToList()];
            }
        }

        private static void SetSearchValidPatternRev(AutogenOptions autogenOptions)
        {
            LocalSpecs.SearchValidPatternRev = string.IsNullOrEmpty(autogenOptions.SearchValidPatternRev) ? LocalSpecs.Options.SearchValidPatternRev : autogenOptions.SearchValidPatternRev.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
            if (LocalSpecs.SearchValidPatternRev)
            {
                LocalSpecs.Options.IsIgnorePatternCheck = false;
            }
        }

        private static void SetSettingAndConfigFiles(AutogenOptions autogenOptions)
        {
            LocalSpecs.SettingFiles = new SettingFiles
            {
                BasicConfigFile = string.IsNullOrEmpty(autogenOptions.BasicConfigFile) ? "" : autogenOptions.BasicConfigFile,
                BinCutInstanceNamingRule = string.IsNullOrEmpty(autogenOptions.BinCutInstanceNamingRule) ? "" : autogenOptions.BinCutInstanceNamingRule,
                RtosCategoryConfig = string.IsNullOrEmpty(autogenOptions.RtosCategoryConfig) ? "" : autogenOptions.RtosCategoryConfig,
                ScanConfig = string.IsNullOrEmpty(autogenOptions.ScanConfig) ? "" : autogenOptions.ScanConfig,
                MbistConfig = string.IsNullOrEmpty(autogenOptions.MbistConfig) ? "" : autogenOptions.MbistConfig,
                SpiConfig = string.IsNullOrEmpty(autogenOptions.SpiConfig) ? "" : autogenOptions.SpiConfig,
                HardipConfig = string.IsNullOrEmpty(autogenOptions.HardipConfig) ? "" : autogenOptions.HardipConfig,
                TnAssignment = string.IsNullOrEmpty(autogenOptions.TnAssignment) ? "" : autogenOptions.TnAssignment,
            };
            if (!string.IsNullOrEmpty(autogenOptions.RepoFolderPath))
            {
                LocalSpecs.SettingFolder = autogenOptions.RepoFolderPath;
            }

            LocalSpecs.ConfigFiles = new ConfigFiles
            {
                BinNumberConfig = string.IsNullOrEmpty(autogenOptions.BinNumberConfig) ? "" : autogenOptions.BinNumberConfig,
            };
            LocalSpecs.FuseCheckFileName = string.IsNullOrEmpty(autogenOptions.FuseCheckTable) ? LocalSpecs.FuseCheckFileName : autogenOptions.FuseCheckTable;
            LocalSpecs.DramTypeFileName = string.IsNullOrEmpty(autogenOptions.DramType) ? LocalSpecs.DramTypeFileName : autogenOptions.DramType;
        }

        private static void ProcessBinCutFiles(AutogenOptions autogenOptions, List<string> selectedFiles)
        {
            if (!string.IsNullOrEmpty(autogenOptions.BinCutFile))
            {
                LocalSpecs.AllBinCutFileNames = [];
                LocalSpecs.BinCutFileName = "";
                foreach (string filePath in autogenOptions.BinCutFile.Split(',').ToList())
                {
                    LocalSpecs.AllBinCutFileNames.Add(filePath);
                }
            }
            foreach (string binCutFileName in LocalSpecs.AllBinCutFileNames)
            {
                string shadowStage = Path.GetFileNameWithoutExtension(binCutFileName).Split('_').Last();
                if (LocalSpecs.AllJobs.Contains(shadowStage))
                {
                    LocalSpecs.BinCutShadowFileNames.Add(binCutFileName);
                }
                else
                {
                    if (string.IsNullOrEmpty(LocalSpecs.BinCutFileName) || LocalSpecs.BinCutFileName.Equals("N/A"))
                    {
                        LocalSpecs.BinCutFileName = binCutFileName;
                    }
                    else
                    {
                        throw new Exception("More than one normal BinCut file has been selected!");
                    }
                }
            }

            if (!string.IsNullOrEmpty(LocalSpecs.BinCutFileName) && !LocalSpecs.BinCutFileName.Equals("N/A", StringComparison.CurrentCultureIgnoreCase))
            {
                selectedFiles.Add(LocalSpecs.BinCutFileName);
            }

            if (LocalSpecs.BinCutShadowFileNames.Count != 0)
            {
                selectedFiles.AddRange(LocalSpecs.BinCutShadowFileNames);
            }
        }

        private static void LoadBinCutWorkbooks()
        {
            if (!string.IsNullOrEmpty(LocalSpecs.BinCutFileName))
            {
                var package = new ExcelPackage(new FileInfo(LocalSpecs.BinCutFileName));
                EpWorkbook.BinCutWorkbook = package.Workbook;
            }

            if (!string.IsNullOrEmpty(LocalSpecs.BinCutPostFileName))
            {
                var package = new ExcelPackage(new FileInfo(LocalSpecs.BinCutPostFileName));
                EpWorkbook.BinCutPostWorkbook = package.Workbook;
            }

            if (!string.IsNullOrEmpty(LocalSpecs.BinCutModeSeqFileName))
            {
                var package = new ExcelPackage(new FileInfo(LocalSpecs.BinCutModeSeqFileName));
                EpWorkbook.BinCutModeSeqWorkbook = package.Workbook;
            }
        }

        private static void ProcessVoltageTables(AutogenOptions autogenOptions, List<string> selectedFiles)
        {
            if (!string.IsNullOrEmpty(autogenOptions.VoltageTable))
            {
                LocalSpecs.VoltageTbFileName = [.. autogenOptions.VoltageTable.Split(',').ToList()];
            }
            selectedFiles.AddRange(LocalSpecs.VoltageTbFileName);
            if (LocalSpecs.VoltageTbFileName.Count != 0)
            {
                MergeSheet.ParseTestSettingSheetToTestPlan(LocalSpecs.VoltageTbFileName, EpWorkbook.TestPlanWorkbook);
            }
        }

        private static void SetByInputFolder(Dictionary<string, string> inputDic)
        {
            if (inputDic.Count != 0)
            {
                LocalSpecs.TestPlanFileName = !inputDic.TryGetValue("Test Plan", out string value) ? "" : value;
                LocalSpecs.ScghFileName = !inputDic.TryGetValue("SCGH", out string value1) ? "" : value1;
                LocalSpecs.PatternListCsvFileName = !inputDic.TryGetValue("Pattern List Csv", out string value2) ? "" : value2;
                if (!inputDic.TryGetValue("Voltage Tables", out string value3))
                {
                    LocalSpecs.VoltageTbFileName = [];
                }
                else
                {
                    LocalSpecs.VoltageTbFileName = [.. value3.Split(',').ToList()];
                }
                if (!inputDic.TryGetValue("Bin Cut", out string value4))
                {
                    LocalSpecs.AllBinCutFileNames = [];
                    LocalSpecs.BinCutFileName = "";
                }
                else
                {
                    LocalSpecs.AllBinCutFileNames = [.. value4.Split(',').ToList()];
                }
                if (!inputDic.TryGetValue("PowerBinning", out string value5))
                {
                    LocalSpecs.AllPowerBinningFileName = [];
                }
                else
                {
                    LocalSpecs.AllPowerBinningFileName = [.. value5.Split(',').ToList()];
                }

                LocalSpecs.BinCutPostFileName = !inputDic.TryGetValue("Post Bin Cut", out string value6) ? "" : value6;
                LocalSpecs.BinCutModeSeqFileName = !inputDic.TryGetValue("BinCut Mode Sequence", out string value7) ? "" : value7;
                LocalSpecs.EquationVoltagesFileName = !inputDic.TryGetValue("Equation Based Voltages", out string value8) ? "N/A" : value8;
                LocalSpecs.TtrSummaryFileName = !inputDic.TryGetValue("HIP TTR Table", out string value9) ? "" : value9;
                LocalSpecs.EfuseTestPlanFileName = !inputDic.TryGetValue("Efuse Test Plan", out string value10) ? "N/A" : value10;
                LocalSpecs.DramTypeFileName = !inputDic.TryGetValue("DRAM Type Table", out string value11) ? "N/A" : value11;
                LocalSpecs.FuseCheckFileName = !inputDic.TryGetValue("Fuse Check Table", out string value12) ? "N/A" : value12;
            }
        }

        private static List<FileInfo> GetVbFiles()
        {
            var mFileList = new List<FileInfo>();
            if (LocalSpecs.Options.Device == EnumDevice.AP)
            {
                mFileList = VbtToolService.GetLibList(FolderStructure.DirLib);
            }
            else if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                mFileList = VbtToolService.GetLibList(FolderStructure.DirLib);
                var dirCodingLib = new DirectoryInfo(FolderStructure.DirConingLib);
                mFileList.AddRange(VbtToolService.GetVbtFilePath(dirCodingLib));
            }
            else if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                mFileList = VbtToolService.GetLibListNonAp(FolderStructure.DirLib);
                mFileList.AddRange(VbtToolService.GetLibList(Path.Combine(FolderStructure.DirLib, "Wireless")));
                VbtToolService.ModifyCommonBas(mFileList);
            }

            return mFileList;
        }
    }
}
