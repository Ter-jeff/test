using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using CommonLib.Static;

using FileDiffLib;

using IgxlLib;
using IgxlLib.Enums;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using LogLib.Static;

using TagDiff.Core.Comparer;
using TagDiff.Core.Output;
using TagDiff.Core.Static;
using TagDiff.Core.Static.ElapsedTime;
using TagDiff.Core.Utility;

namespace TagDiff.Core
{
    public class TagDiffMain
    {
        private readonly string _autoGenTestProgramPath;
        private readonly string _productionTestProgramPath;
        private readonly string _outputPath;
        private readonly bool _onlyCompareCsharpInstance;
        private readonly string _job;
        private readonly bool _isUnitTest;
        private readonly string _reportPath;
        private readonly IEnumerable<string> _excludeDirs;

        private readonly Dictionary<string, List<string>> _fileInfos;
        private readonly HashSet<string> _skipFiles = new(
        [
            "Execution%20Profile.txt",
            "_NonSheetReferences_.txt",
            "ThisWorkbook.txt",
            "tl_WorkBookProperties_.txt",
            "_NonSheetReferences_.txt",
            "_RangeNames_.txt",
            "IGLinkManifest.txt",
            "InputFilesInfo.txt"
        ], StringExtensions.IgnoreCase);

        public static readonly bool IncludingBackup = false;

        public TagDiffMain(string autoGenTestProgramPath, string productionTestProgramPath, string outputPath, bool onlyCompareCsharpInstance, string job, bool isUnitTest = false, IEnumerable<string>? excludeDirs = null)
        {
            _autoGenTestProgramPath = Path.GetFullPath(autoGenTestProgramPath);
            _productionTestProgramPath = Path.GetFullPath(productionTestProgramPath);
            _outputPath = Path.GetFullPath(outputPath);
            _onlyCompareCsharpInstance = onlyCompareCsharpInstance;
            _job = job;
            _isUnitTest = isUnitTest;
            _excludeDirs = excludeDirs ?? [];
            string name = Path.GetFileNameWithoutExtension(_autoGenTestProgramPath);
            _reportPath = Path.Combine(_outputPath, $"{name}_AutoRate_{TimeContext.Now:yyMMddHHmm}.xlsx");

            _fileInfos = new Dictionary<string, List<string>>
            {
                { "Production Program", [_productionTestProgramPath] },
                { "Autogen Program", [_autoGenTestProgramPath] }
            };
            var assembly = Assembly.GetExecutingAssembly();
            string fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version ?? assembly.GetName().Version?.ToString() ?? "0.0.0.0";
            string version = AssemblyProvider.Current.GetFileVersion(fileVersion);
            string assemblyName = assembly.GetName().Name ?? "";
            _fileInfos.Add("Version", [$"{assemblyName}_{version}"]);
        }

        public void WorkFlow()
        {
            try
            {
                TagDiffStatic.ReturnValue = 0;
                ErrorReportManager.ClearErrors();

                string outputPath = Path.Combine(_outputPath, "Production");
                IgxlExport.ExportFiles(_productionTestProgramPath, outputPath);
                string expectPath = Path.Combine(_outputPath, "Autogen");
                IgxlExport.ExportFiles(_autoGenTestProgramPath, expectPath);

                string batFolder = Path.Combine(_outputPath, "Diff");
                string oasisRootFolder = Environment.GetEnvironmentVariable("OASISROOT") ?? "";
                if (!string.IsNullOrEmpty(oasisRootFolder))
                {
                    string absoluteBatFolder = Path.GetFullPath(batFolder);
                    string absoluteOutputPath = Path.GetFullPath(outputPath);
                    string absoluteExpectPath = Path.GetFullPath(expectPath);

                    new FileComparisonReport("TagDiff").GenDiffBats(absoluteBatFolder, absoluteOutputPath, absoluteExpectPath, oasisRootFolder);
                }

                ConcurrentDictionary<string, (EnumSheetType sheetType, string baseFile, string compareFile)> fileMappings = GetFileMappings(outputPath, expectPath, _skipFiles);

                List<CompareReport> allCompareReports = CompareSheet(fileMappings, _productionTestProgramPath, _job, _onlyCompareCsharpInstance);

                WriteReport(_outputPath, batFolder, fileMappings, _reportPath, allCompareReports, _isUnitTest, _fileInfos, _excludeDirs);

                IgxlExport.ExportFiles(_productionTestProgramPath, outputPath);
            }
            catch (Exception e)
            {
                Response.Report(e.Message, EnumMessageLevel.Error, 0);
                Response.Report(e.StackTrace ?? "No stack trace available.", EnumMessageLevel.Error, 0);
                TagDiffStatic.ReturnValue = 1;
            }
        }

        private static ConcurrentDictionary<string, (EnumSheetType sheetType, string baseFile, string compareFile)> GetFileMappings(string outputPath, string expectPath, HashSet<string> skipFiles)
        {
            string[] outputFiles = Directory.GetFiles(outputPath, "*.*", SearchOption.TopDirectoryOnly);
            string[] expectedFiles = Directory.GetFiles(expectPath, "*.*", SearchOption.TopDirectoryOnly);

            var fileMappings = new ConcurrentDictionary<string, (EnumSheetType sheetType, string baseFile, string compareFile)>();
            //foreach (string outputFile in outputFiles)
            Parallel.ForEach(outputFiles, outputFile =>
            {
                string fileName = Path.GetFileName(outputFile);
                string ext = Path.GetExtension(outputFile);
                if (!ext.ContainsIgnoreCase(".txt") || skipFiles.Contains(fileName))
                {
                    //continue;
                    return;
                }

                EnumSheetType sheetType = GetIgxlSheetTypeByFile(outputFile);
                bool isJobList = sheetType == EnumSheetType.DTJobListSheet;
                string expectedFile = expectedFiles.FirstOrDefault(f => isJobList ? GetIgxlSheetTypeByFile(f) == EnumSheetType.DTJobListSheet : Path.GetFileName(f).EqualsIgnoreCase(fileName)) ?? "";
                string sheetName = Path.GetFileNameWithoutExtension(fileName);
                fileMappings[sheetName] = (sheetType, outputFile, expectedFile);
            }
        );
            return fileMappings;
        }

        public static EnumSheetType GetIgxlSheetTypeByFile(string file)
        {
            string? text;
            using (var streamReader = new StreamReader(file))
            {
                text = streamReader.ReadLine();
            }

            string extension = Path.GetExtension(file);
            if (extension.EqualsIgnoreCase(".bas"))
            {
                return EnumSheetType.Bas;
            }

            if (extension.EqualsIgnoreCase(".cls"))
            {
                return EnumSheetType.Cls;
            }

            return IgxlLoaderHelpers.GetIgxlSheetType(text);
        }

        private static List<CompareReport> CompareSheet(ConcurrentDictionary<string, (EnumSheetType sheetType, string baseFile, string compareFile)> fileMappings, string productionTestProgramPath, string job, bool onlyCompareCsharpInstance)
        {
            Response.Report("Loading Test Program ...", EnumMessageLevel.Info, 10);
            var types = new List<EnumSheetType> { EnumSheetType.DTFlowtableSheet, EnumSheetType.DTTestInstancesSheet };
            var igxlDataLoader = new IgxlLoader(productionTestProgramPath, types);
            List<SubFlowSheet> flowSheets = igxlDataLoader.FlowSheets;
            List<InstanceSheet> instanceSheets = igxlDataLoader.InstanceSheets;
            var instance = flowSheets.SelectMany(x => x.Rows).Where(x => x.Opcode.EqualsIgnoreCase("Test")).Select(x => x.Parameter).ToList();
            var instanceRows = instanceSheets.SelectMany(x => x.Rows).Where(x => instance.Exists(y => y.EqualsIgnoreCase(x.TestName))).ToList();

            var compareReports = new ConcurrentBag<CompareReport>();
            var allCompareReports = new List<CompareReport>();
            ElapsedTimeStatic.Clear();
            int totalItems = fileMappings.Count;
            int completedItems = 0;
            Parallel.ForEach(fileMappings, fileTuple =>
            {
                int currentCompleted = Interlocked.Increment(ref completedItems);
                int percentage = currentCompleted * 100 / totalItems;
                ElapsedTimeItem elapsedTimeItem = StopwatchMeasureSection.MeasureSection(fileTuple.Key, () =>
                {
                    string sheetName = fileTuple.Key;
                    Response.Report($"Comparing {sheetName} sheet ...", EnumMessageLevel.General, percentage);
                    SheetComparerBase? igxlSheetComparer = new SheetComparerFactory(onlyCompareCsharpInstance).GetIgxlSheetComparer(sheetName, fileTuple.Value.sheetType, flowSheets, instanceRows, job, fileTuple.Value.baseFile, fileTuple.Value.compareFile);
                    if (igxlSheetComparer != null)
                    {
                        List<CompareReport> compResults = igxlSheetComparer.Compare();
                        if (compResults.Count != 0)
                        {
                            for (int i = 0; i < compResults.Count; i++)
                            {
                                CompareReport compareReport = compResults[i];
                                ClassifyCompareReport.SetCategory(fileTuple.Value.sheetType, sheetName, ref compareReport);
                                compareReports.Add(compareReport);
                            }
                        }
                    }
                });

                ElapsedTimeStatic.Add(elapsedTimeItem);
            }
            );

            allCompareReports.AddRange(compareReports);
            return allCompareReports;
        }

        private static void WriteReport(string outputPath, string batFolder, ConcurrentDictionary<string, (EnumSheetType sheetType, string baseFile, string compareFile)> fileMappings, string reportPath, List<CompareReport> compareReports, bool isUnitTest, Dictionary<string, List<string>> fileInfos, IEnumerable<string> excludeDirs)
        {
            Response.Report("Starting to Write Report ...", EnumMessageLevel.General, 90);
            var reportWriter = new ReportWriter(isUnitTest, outputPath, fileInfos);
            reportWriter.WorkFlow(batFolder, fileMappings, reportPath, compareReports, excludeDirs);

            Response.Report("Printing Elapsed Time ...", EnumMessageLevel.General, 90);
            ElapsedTimeStatic.WriteToJson(outputPath, ElapsedTimeStatic.ElapsedTimeItems, compareReports);

            Response.Report(reportPath, EnumMessageLevel.General, 90);
            Response.Report("Completed!", EnumMessageLevel.EndPoint, 100);
        }
    }
}
