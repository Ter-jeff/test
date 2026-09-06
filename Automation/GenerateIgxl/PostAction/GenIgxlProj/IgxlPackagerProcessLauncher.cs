using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

using Automation.IgxlPackaging.Contract;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using Newtonsoft.Json;

using TestPlanLib.Basic;

namespace Automation.GenerateIgxl.PostAction.GenIgxlProj
{
    /// <summary>
    /// Replaces the in-process IgxlProjMain. Builds a JSON descriptor of the IGLink projects
    /// to generate and shells out to Automation.IgxlPackaging.exe (net48 + IGLinkBase).
    /// </summary>
    internal sealed class IgxlPackagerProcessLauncher : IIgxlPackager
    {
        private const string CommonCodeAndSheetsName = "CommonCodeAndSheets";

        private readonly string _packagerExe;
        private readonly string _monoExe;

        public IgxlPackagerProcessLauncher(string packagerExe, string monoExe)
        {
            _packagerExe = packagerExe;
            _monoExe = monoExe;
        }

        public void GenIgxlProg(List<string> sourceFile, string outputFolder, string programName, IgxlWorkBook igxlWorkBook, bool isUnitTest)
        {
            try
            {
                string tmpIgLinkFolder = Path.Combine(outputFolder, "IGLink");
                if (!Directory.Exists(tmpIgLinkFolder))
                {
                    Directory.CreateDirectory(tmpIgLinkFolder);
                }

                List<string> filteredSources = DedupeSources(sourceFile);

                var request = new PackageRequest();
                if (TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any())
                {
                    foreach (TestProgramRow row in TestPlanStatic.TestProgramDefSheet.Rows)
                    {
                        if (row.JobMapping.Any())
                        {
                            DeviceProjectDescriptor descriptor = BuildDescriptor(filteredSources, row.ProgramName, igxlWorkBook, tmpIgLinkFolder);
                            if (descriptor != null)
                            {
                                request.Projects.Add(descriptor);
                            }
                        }
                    }
                }
                else
                {
                    DeviceProjectDescriptor descriptor = BuildDescriptor(filteredSources, programName, igxlWorkBook, tmpIgLinkFolder);
                    if (descriptor != null)
                    {
                        request.Projects.Add(descriptor);
                    }
                }

                if (!request.Projects.Any())
                {
                    Response.Report("IGXL Auto Gen is Completed !", percentage: 100);
                    return;
                }

                InvokeSidecar(request);

                if (!isUnitTest)
                {
                    var igxlManager = new IgxlManager();
                    foreach (DeviceProjectDescriptor project in request.Projects)
                    {
                        Response.Report("Creating IGXL workbook ...", percentage: 80);
                        string projFile = project.OutputProjFile;
                        string subProgramName = Path.GetFileNameWithoutExtension(projFile);
                        string outputIgxl = Path.ChangeExtension(projFile, ".igxl");
                        igxlManager.GenTestProgramBySubProgram(projFile, outputIgxl, " -g ", subProgramName);
                    }
                }

                Response.Report("IGXL Auto Gen is Completed !", percentage: 100);
            }
            catch (Exception e)
            {
                Response.Report($"Error occurs during generate IGlink Project: {e.GetType().FullName}: {e.Message}\n{e}", EnumMessageLevel.Error, 0);
            }
        }

        private static List<string> DedupeSources(List<string> sourceFile)
        {
            var filterSourceFile = new Dictionary<string, string>();
            foreach (string data in sourceFile)
            {
                string fileName = Path.GetFileName(data).ToLower();
                if (!filterSourceFile.ContainsKey(fileName))
                {
                    filterSourceFile.Add(fileName, data);
                }
                else
                {
                    ErrorReportManager.AddError(PostActionErrorType.W_DuplicateFile_01, "", 1, 0, [fileName]);
                    filterSourceFile[fileName] = data;
                }
            }
            return filterSourceFile.Values.ToList();
        }

        internal static DeviceProjectDescriptor BuildDescriptor(List<string> filteredSources, string programName, IgxlWorkBook igxlWorkBook, string tmpIgLinkFolder)
        {
            bool isSubprogram = programName.EndsWith("_Sub");
            bool hasTestProgramDef = TestPlanStatic.TestProgramDefSheet != null && TestPlanStatic.TestProgramDefSheet.Rows.Any();

            Dictionary<string, MainFlow> allMainFlows = igxlWorkBook.MainFlowSheets;
            var byPassJobLists = new List<string>();
            if (hasTestProgramDef)
            {
                ApplyTestProgramDefScope(programName, igxlWorkBook, ref allMainFlows, ref byPassJobLists);
            }

            string defaultChannelMap = LocalSpecs.DefaultChannelMap == null
                ? string.Empty
                : igxlWorkBook.ChannelMapSheets.Values.ToList().Exists(x => x.Name.Equals(LocalSpecs.DefaultChannelMap))
                    ? LocalSpecs.DefaultChannelMap
                    : string.Empty;

            var subProgramTotal = new SubProgramDescriptor
            {
                Name = programName,
                JobNames = programName,
                GenerateJobListSheet = false,
            };
            var commonCodeSheets = new List<string>();
            var commonCodeVbFiles = new List<string>();
            var folderSubPrograms = new List<SubProgramDescriptor>();
            var subProgramSubs = new List<SubProgramDescriptor>();

            if (isSubprogram)
            {
                foreach (SubprogramSetting setting in TestPlanStatic.SubprogramMappingSheet.SubprogramSettings)
                {
                    subProgramSubs.Add(new SubProgramDescriptor
                    {
                        Name = setting.SubprogramName,
                        JobNames = setting.SubprogramName,
                    });
                }
            }

            Response.Report("Generating IGXL Test Program ...", percentage: 0);

            foreach (string data in filteredSources)
            {
                AddSourceToPackage(data, tmpIgLinkFolder, isSubprogram, byPassJobLists, subProgramTotal,
                    commonCodeSheets, commonCodeVbFiles, subProgramSubs, folderSubPrograms);
            }

            string mainFlow = allMainFlows.Any()
                ? programName + ":" + allMainFlows.First().Value.Name
                : programName + ":";

            var allJobs = new List<string>();
            if (allMainFlows.Any())
            {
                foreach (SubFlowSheet main in allMainFlows.Select(x => x.Value))
                {
                    if (main?.JobNames != null)
                    {
                        allJobs.AddRange(main.JobNames);
                    }
                }
            }

            subProgramTotal.MainFlow = mainFlow;

            var descriptor = new DeviceProjectDescriptor
            {
                Name = "ProjectTemple",
                OutputProjFile = Path.Combine(tmpIgLinkFolder, programName + ".igxlProj"),
                SaveAsXLS = false,
                SheetOrder = "Alphabetically",
                CurrentProject = LocalSpecs.CurrentProject,
                DefaultChannelMap = defaultChannelMap,
                CommonCode = new CommonCodeDescriptor
                {
                    SheetSources = commonCodeSheets,
                    VbFileSources = commonCodeVbFiles,
                },
            };
            descriptor.SubPrograms.Add(subProgramTotal);

            foreach (SubProgramDescriptor sub in folderSubPrograms)
            {
                sub.MainFlow = mainFlow;
                descriptor.SubPrograms.Add(sub);
            }

            if (subProgramSubs.Any())
            {
                foreach (SubProgramDescriptor sub in subProgramSubs)
                {
                    string mainflowSheetName = sub.SheetSources
                        .Select(Path.GetFileNameWithoutExtension)
                        .FirstOrDefault(name => name != null && name.StartsWith("Main_", StringComparison.OrdinalIgnoreCase));
                    sub.MainFlow = $"{sub.Name}:{mainflowSheetName}";
                    sub.GenerateJobListSheet = false;
                    sub.FlowGenMode = "UseMainFlowOnly";
                    descriptor.SubPrograms.Add(sub);
                }
            }

            foreach (string job in allJobs)
            {
                string mappingJob = job;
                if (hasTestProgramDef)
                {
                    Dictionary<string, string> jobMapping = TestPlanStatic.TestProgramDefSheet.Rows
                        .Find(x => x.ProgramName.Equals(programName)).JobMapping;
                    mappingJob = jobMapping[job];
                }
                descriptor.Jobs.Add(BuildJobDescriptor(job, "Main_Flow_" + mappingJob, allJobs, folderSubPrograms, defaultChannelMap));
                descriptor.WorkBooks.Add(BuildWorkbookDescriptor(descriptor.Jobs, defaultChannelMap, job));
            }

            Response.Report("Saving IGlink Project ...", percentage: 40);
            return descriptor;
        }

        private static void ApplyTestProgramDefScope(string projectName, IgxlWorkBook igxlWorkBook,
            ref Dictionary<string, MainFlow> allMainFlows, ref List<string> byPassJobLists)
        {
            TestProgramRow jobMappingRow = TestPlanStatic.TestProgramDefSheet.Rows.Find(x => x.ProgramName == projectName);
            byPassJobLists = TestProgram.IgxlWorkBk.JobListSheets
                .Where(x => !x.Value.Name.Equals(jobMappingRow.JobListSheet.Name))
                .Select(x => x.Key + ".txt")
                .ToList();
            allMainFlows = new Dictionary<string, MainFlow>();
            foreach (string job in jobMappingRow.JobMapping.Values)
            {
                string mainFlowName = $"Main_Flow_{job}";
                KeyValuePair<string, MainFlow> mainFlowByJob = igxlWorkBook.MainFlowSheets.ToList().Find(x => x.Value.Name == mainFlowName);
                allMainFlows.Add(mainFlowByJob.Key, mainFlowByJob.Value);
            }
            foreach (string mainflow in igxlWorkBook.MainFlowSheets.Keys)
            {
                if (allMainFlows.ContainsKey(mainflow))
                {
                    continue;
                }
                byPassJobLists.Add(mainflow + ".txt");
            }
        }

        private static void AddSourceToPackage(string data, string tmpIgLinkFolder, bool isSubprogram, List<string> byPassJobLists,
            SubProgramDescriptor subProgramTotal, List<string> commonCodeSheets, List<string> commonCodeVbFiles,
            List<SubProgramDescriptor> subProgramSubs, List<SubProgramDescriptor> folderSubPrograms)
        {
            if (byPassJobLists.Exists(x => x.ToLower().Equals(data.ToLower())))
            {
                return;
            }

            if (data.IndexOf("VBT_Instrument_Setup", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return;
            }

            if (data.IndexOf(".txt", StringComparison.OrdinalIgnoreCase) != -1)
            {
                string sheetSource = MakeRelative(data, tmpIgLinkFolder);
                subProgramTotal.SheetSources.Add(sheetSource);

                if (isSubprogram)
                {
                    string sheetName = Path.GetFileNameWithoutExtension(data);
                    if (sheetName.StartsWith("Main_", StringComparison.CurrentCultureIgnoreCase) ||
                        sheetName.StartsWith("JobList_", StringComparison.CurrentCultureIgnoreCase))
                    {
                        SubProgramDescriptor target = subProgramSubs.Find(x => sheetName.EndsWith(x.Name));
                        target?.SheetSources.Add(sheetSource);
                    }
                    else if (StartsWithStaticDir(data))
                    {
                        commonCodeSheets.Add(sheetSource);
                    }
                    else
                    {
                        foreach (SubProgramDescriptor sub in subProgramSubs)
                        {
                            sub.SheetSources.Add(sheetSource);
                        }
                    }
                }
                else
                {
                    if (StartsWithStaticDir(data))
                    {
                        commonCodeSheets.Add(sheetSource);
                    }
                    else
                    {
                        string folderName = Path.GetFileName(Path.GetDirectoryName(data));
                        if (!string.IsNullOrEmpty(folderName))
                        {
                            SubProgramDescriptor existing = folderSubPrograms.Find(x => x.Name.Equals(folderName, StringComparison.CurrentCultureIgnoreCase));
                            if (existing != null)
                            {
                                existing.SheetSources.Add(sheetSource);
                            }
                            else
                            {
                                folderSubPrograms.Add(new SubProgramDescriptor
                                {
                                    Name = folderName,
                                    JobNames = folderName,
                                    SheetSources = { sheetSource },
                                });
                            }
                        }
                    }
                }
            }
            else if (IsVbFile(data))
            {
                string vbSource = MakeRelative(data, tmpIgLinkFolder);
                subProgramTotal.VbFileSources.Add(vbSource);
                commonCodeVbFiles.Add(vbSource);
            }
        }

        internal static DeviceJobDescriptor BuildJobDescriptor(string jobName, string mainFlow, List<string> allJobs, List<SubProgramDescriptor> subPrograms, string channelMap)
        {
            var job = new DeviceJobDescriptor
            {
                Name = jobName,
                GenerateJobListSheet = false,
                GenerateExecIPModule = false,
                FlowGenMode = "UseMainFlowOnly",
                MainFlow = $"{CommonCodeAndSheetsName}: {mainFlow}",
                AppendToFlow = false,
                PinMap = $"{CommonCodeAndSheetsName}: PinMap",
                JobNames = string.Join(",", allJobs),
                DefaultJob = jobName,
                ChannelMapDisplayMode = "Signal",
                IncludeOnlyOnePinMap = false,
                IncludeOnlyDefaultChanMap = false,
            };
            job.SubProgramNames.AddRange(subPrograms.Select(x => x.Name));
            if (!string.IsNullOrEmpty(channelMap))
            {
                job.DefaultChannelMap = $"{CommonCodeAndSheetsName}: {channelMap}";
            }
            return job;
        }

        internal static WorkbookDescriptor BuildWorkbookDescriptor(List<DeviceJobDescriptor> jobs, string channelMap, string defaultJob)
        {
            var workbook = new WorkbookDescriptor
            {
                Name = LocalSpecs.CurrentProject,
                GenerateJobListSheet = false,
                ChannelMapDisplayMode = "Signal",
                DefaultJob = defaultJob,
            };
            if (!string.IsNullOrEmpty(channelMap))
            {
                workbook.DefaultChannelMap = $"{CommonCodeAndSheetsName}: {channelMap}";
            }
            workbook.Jobs.AddRange(jobs.Select(j => j.Name));
            return workbook;
        }

        internal static bool IsVbFile(string path)
        {
            return path.IndexOf(".bas", StringComparison.OrdinalIgnoreCase) != -1
                || path.IndexOf(".cls", StringComparison.OrdinalIgnoreCase) != -1
                || path.IndexOf(".frm", StringComparison.OrdinalIgnoreCase) != -1
                || path.IndexOf(".frx", StringComparison.OrdinalIgnoreCase) != -1;
        }

        internal static bool StartsWithStaticDir(string path)
        {
            return path.StartsWith(FolderStructure.DirCommon, StringComparison.CurrentCultureIgnoreCase)
                || path.StartsWith(FolderStructure.DirLib, StringComparison.CurrentCultureIgnoreCase)
                || path.StartsWith(FolderStructure.DirMain, StringComparison.CurrentCultureIgnoreCase);
        }

        /// <summary>
        /// IGLinkBase's <c>Sheet.Source</c> historically gets the path relative to the IGLink folder
        /// (when the path is under it) or the full path otherwise — this preserves that behavior.
        /// </summary>
        internal static string MakeRelative(string filepath, string refPath)
        {
            if (string.IsNullOrEmpty(refPath))
            {
                return filepath;
            }
            if (filepath.IndexOf(refPath, StringComparison.OrdinalIgnoreCase) != -1)
            {
                return filepath.Substring(refPath.Length + 1, filepath.Length - refPath.Length - 1);
            }
            return filepath;
        }

        private void InvokeSidecar(PackageRequest request)
        {
            string descriptorPath = Path.Combine(Path.GetTempPath(), $"igxl-package-{Guid.NewGuid():N}.json");
            File.WriteAllText(descriptorPath, JsonConvert.SerializeObject(request, Formatting.Indented));

            try
            {
                var psi = new ProcessStartInfo
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                };
                if (_monoExe != null)
                {
                    psi.FileName = _monoExe;
                    psi.Arguments = $"\"{_packagerExe}\" \"{descriptorPath}\"";
                }
                else
                {
                    psi.FileName = _packagerExe;
                    psi.Arguments = $"\"{descriptorPath}\"";
                }

                Response.Report($"Invoking IGXL packager: {psi.FileName} {psi.Arguments}");

                using (var process = new Process { StartInfo = psi })
                {
                    process.OutputDataReceived += (_, e) => { if (e.Data != null) { Response.Report(e.Data); } };
                    process.ErrorDataReceived += (_, e) => { if (e.Data != null) { Response.Report(e.Data, EnumMessageLevel.Error); } };
                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        Response.Report($"IGXL packager exited with code {process.ExitCode}", EnumMessageLevel.Error);
                    }
                }
            }
            finally
            {
                try
                {
                    File.Delete(descriptorPath);
                }
                catch
                {
                    /* best-effort cleanup */
                }
            }
        }
    }
}
