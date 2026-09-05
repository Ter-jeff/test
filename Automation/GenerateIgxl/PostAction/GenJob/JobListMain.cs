using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using TestPlanLib.Basic;

namespace Automation.GenerateIgxl.PostAction.GenJob
{
    public class JobListMain
    {
        private const string JobListName = "JobList";
        private const string T0TxJobListName = "JobList_T0Tx";

        public void WorkFlow(TestProgramRow testProgramDefRow = null)
        {
            string patternGroups = "";
            string testProcedures = "";
            string signals = "";
            string fractionalBus = "";
            string comment = "";
            var portMapList = TestProgram.IgxlWorkBk.PortMapSheets.Where(x => x.Value.Rows.Count > 0).Select(x => x.Value.Name).OrderBy(x => x).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (LocalSpecs.JobMap.Count == 0)
            {
                return;
            }

            #region Edit Item
            var insSheets = TestProgram.IgxlWorkBk.InsSheets.Select(x => x.Value.Name).ToList();
            foreach (string customPath in LocalSpecs.CustomPath)
            {
                if (Directory.Exists(customPath))
                {
                    List<string> extraInst = IgxlSheetReaderHelpers.GetSheetsByType(customPath, EnumSheetType.DTTestInstancesSheet);
                    insSheets.AddRange(extraInst.Select(Path.GetFileNameWithoutExtension));
                }
            }
            string acSpecs = string.Join(",", TestProgram.IgxlWorkBk.AcSpecSheets.Where(x => x.Value.CategoryList.Count > 0).Select(x => x.Value.Name).OrderBy(x => x));
            string.Join(",", TestProgram.IgxlWorkBk.DcSpecSheets.Where(x => x.Value.CategoryList.Count > 0).Select(x => x.Value.Name).OrderBy(x => x));
            string patternSets = string.Join(",", TestProgram.IgxlWorkBk.PatSetSheets.Where(x => !x.Value.Name.Equals("PatSets_DashBoard") && x.Value.Rows.Any()).Select(x => x.Value.Name).OrderBy(x => x));
            string binTable = string.Join(",", TestProgram.IgxlWorkBk.BinTblSheets.Where(x => x.Value.Rows.Count > 0).Select(x => x.Value.Name).OrderBy(x => x));
            string portMap = string.Join(",", portMapList);
            string characterization = string.Join(",", TestProgram.IgxlWorkBk.CharSheets.Select(x => x.Value.Name).OrderBy(x => x));
            string mixedSignalTiming = string.Join(",", TestProgram.IgxlWorkBk.MixedSignalSheets.Select(x => x.Value.Name).OrderBy(x => x));
            string waveDefinition = string.Join(",", TestProgram.IgxlWorkBk.WaveDefSheets.Select(x => x.Value.Name).OrderBy(x => x));
            #endregion

            var context = new JobRowContext
            {
                InsSheets = insSheets,
                PortMapList = portMapList,
                AcSpecs = acSpecs,
                PatternSets = patternSets,
                PatternGroups = patternGroups,
                BinTable = binTable,
                Characterization = characterization,
                TestProcedures = testProcedures,
                MixedSignalTiming = mixedSignalTiming,
                WaveDefinition = waveDefinition,
                Signals = signals,
                FractionalBus = fractionalBus,
                Comment = comment
            };

            List<string> jobList = LocalSpecs.AllJobs;
            #region Generate JobList for main program
            var jobListSheet = new JobListSheet(JobListName);
            if (testProgramDefRow != null)
            {
                jobList = testProgramDefRow.JobMapping.Keys.ToList();
                jobListSheet = new JobListSheet($"{JobListName}_{testProgramDefRow.ProgramName}");
            }
            foreach (string jobName in jobList)
            {
                JobRow jobRow = CreateJobRow(context, jobName, ref portMap);
                if (testProgramDefRow != null)
                {
                    string jobMappingName = testProgramDefRow.JobMapping[jobName];
                    jobRow.FlowTable = $"Main_Flow_{jobMappingName}";
                }
                else
                {
                    jobRow.FlowTable = FindJobMainFlow(jobName);
                }

                jobRow.DcSpecs = FindDcSpec(jobName);
                SetConcurrentSequence(jobRow, jobName.ToUpper(), jobName);
                jobListSheet.AddRow(jobRow);
            }
            AddJobListSheetForMain(testProgramDefRow, jobListSheet);
            #endregion

            #region Generate JobList for subprogram
            if (TestProgram.SubProgIgxlWorkBk.MainFlowSheets.Any())
            {
                foreach (KeyValuePair<string, MainFlow> subprogramMainFlow in TestProgram.SubProgIgxlWorkBk.MainFlowSheets)
                {
                    var subJobListSheet = new JobListSheet($"{JobListName}_{subprogramMainFlow.Value.Name.Replace("Main_Flow_", "")}");
                    if (testProgramDefRow != null)
                    {
                        jobList = testProgramDefRow.JobMapping.Keys.ToList();
                        subJobListSheet = new JobListSheet($"{JobListName}_{subprogramMainFlow.Value.Name.Replace("Main_Flow_", "")}_{testProgramDefRow.ProgramName}");
                    }
                    foreach (string jobName in jobList)
                    {
                        JobRow jobRow = CreateJobRow(context, jobName, ref portMap);
                        jobRow.FlowTable = subprogramMainFlow.Value.Name;
                        jobRow.DcSpecs = FindDcSpec(jobName);
                        SetConcurrentSequence(jobRow, jobName.ToUpper(), jobName);
                        subJobListSheet.AddRow(jobRow);
                    }
                    AddJobListSheetForSubProgram(testProgramDefRow, subJobListSheet);
                }
            }
            #endregion

            #region Generate JobList for T0TX program
            if (LocalSpecs.Options.GenerateT0TXTestprogram)
            {
                var subJobListSheet = new JobListSheet(T0TxJobListName);
                var jobListT0Tx = new List<string>() { "FT1", "FT2" };
                if (testProgramDefRow != null)
                {
                    jobList = testProgramDefRow.JobMapping.Keys.ToList();
                    subJobListSheet = new JobListSheet($"{T0TxJobListName}_{testProgramDefRow.ProgramName}");
                }
                foreach (string jobName in jobListT0Tx)
                {
                    JobRow jobRow = CreateJobRow(context, jobName, ref portMap);
                    KeyValuePair<string, MainFlow> targetMainFlow = jobName.Equals("FT1") ? TestProgram.T0TxIgxlWorkBk.MainFlowSheets.FirstOrDefault(x => x.Value.Name.ContainsIgnoreCase("Room")) : TestProgram.T0TxIgxlWorkBk.MainFlowSheets.FirstOrDefault(x => x.Value.Name.ContainsIgnoreCase("Hot"));
                    jobRow.FlowTable = targetMainFlow.Value.Name;
                    jobRow.DcSpecs = FindDcSpec(jobName);
                    SetConcurrentSequence(jobRow, "T0TX_Room", jobName);
                    subJobListSheet.AddRow(jobRow);
                }
                AddJobListSheetForT0Tx(testProgramDefRow, subJobListSheet);
            }
            #endregion
        }

        private class JobRowContext
        {
            public List<string> InsSheets;
            public HashSet<string> PortMapList;
            public string AcSpecs;
            public string PatternSets;
            public string PatternGroups;
            public string BinTable;
            public string Characterization;
            public string TestProcedures;
            public string MixedSignalTiming;
            public string WaveDefinition;
            public string Signals;
            public string FractionalBus;
            public string Comment;
        }

        private static JobRow CreateJobRow(JobRowContext context, string jobName, ref string portMap)
        {
            var jobRow = new JobRow
            {
                JobName = jobName,
                PinMap = TestProgram.IgxlWorkBk.PinMapPair.Value == null
                    ? ""
                    : TestProgram.IgxlWorkBk.PinMapPair.Value.Name,
                TestInstances = string.Join(",", context.InsSheets.OrderBy(x => x))
            };
            jobRow.AcSpecs = context.AcSpecs;
            jobRow.PatternSets = context.PatternSets;
            jobRow.PatternGroups = context.PatternGroups;
            jobRow.BinTable = context.BinTable;
            jobRow.Characterization = context.Characterization;
            jobRow.TestProcedures = context.TestProcedures;
            jobRow.MixedSignalTiming = context.MixedSignalTiming;
            jobRow.WaveDefinitions = context.WaveDefinition;
            jobRow.Signals = context.Signals;
            ApplyPortMap(context.PortMapList, jobRow, jobName, ref portMap);
            jobRow.FractionalBus = context.FractionalBus;
            jobRow.Comment = context.Comment;
            return jobRow;
        }

        private static void ApplyPortMap(HashSet<string> portMapList, JobRow jobRow, string jobName, ref string portMap)
        {
            if (TestPlanStatic.JobInfoSheet != null && TestPlanStatic.JobInfoSheet.Rows.Find(x => x.JobName.Equals(jobName)) != null && !string.IsNullOrEmpty(portMap))
            {
                string portMapName = TestPlanStatic.JobInfoSheet.Rows.Find(x => x.JobName.Equals(jobName)).TesterType.ToUpper() == "UF" ? "PortMap_UF" : "PortMap_UFP";
                if (portMapList.Contains(portMapName))
                {
                    portMap = portMapName;
                }
            }
            if (!string.IsNullOrEmpty(portMap))
            {
                jobRow.PortMap = portMap;
            }
        }

        private static void SetConcurrentSequence(JobRow jobRow, string anyKey, string jobName)
        {
            if (TestProgram.IgxlWorkBk.MainFlowSheets.Any(x => x.Value.Name.ContainsIgnoreCase(anyKey)))
            {
                jobRow.ConcurrentSequence =
                    TestProgram.IgxlWorkBk.MainFlowSheets.First(
                        x => x.Value.Name.ContainsIgnoreCase(jobName.ToUpper())).Value.HasConcurrent
                        ? "Concurrent Sequence"
                        : "";
            }
            else
            {
                jobRow.ConcurrentSequence = "";
            }
        }

        private static void AddJobListSheetForMain(TestProgramRow testProgramDefRow, JobListSheet jobListSheet)
        {
            if (testProgramDefRow == null)
            {
                TestProgram.IgxlWorkBk.AddJobListSheet(FolderStructure.DirJob, jobListSheet);
            }
            else
            {
                TestProgram.IgxlWorkBk.AddJobListSheet(FolderStructure.DirJob, jobListSheet);
                testProgramDefRow.JobListSheet = jobListSheet;
            }
        }

        private static void AddJobListSheetForSubProgram(TestProgramRow testProgramDefRow, JobListSheet subJobListSheet)
        {
            if (testProgramDefRow == null)
            {
                TestProgram.SubProgIgxlWorkBk.AddJobListSheet(FolderStructure.DirSubProgram, subJobListSheet);
            }
            else
            {
                TestProgram.SubProgIgxlWorkBk.AddJobListSheet(FolderStructure.DirSubProgram, subJobListSheet);
                testProgramDefRow.JobListSheet = subJobListSheet;
            }
        }

        private static void AddJobListSheetForT0Tx(TestProgramRow testProgramDefRow, JobListSheet subJobListSheet)
        {
            if (testProgramDefRow == null)
            {
                TestProgram.T0TxIgxlWorkBk.AddJobListSheet(FolderStructure.DirT0Tx, subJobListSheet);
            }
            else
            {
                TestProgram.T0TxIgxlWorkBk.AddJobListSheet(FolderStructure.DirT0Tx, subJobListSheet);
                testProgramDefRow.JobListSheet = subJobListSheet;
            }
        }

        private string FindJobMainFlow(string job)
        {
            foreach (KeyValuePair<string, MainFlow> mainFlow in TestProgram.IgxlWorkBk.MainFlowSheets)
            {
                string name = mainFlow.Value.Name;
                if (name.IndexOf(job, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return name;
                }
                if (mainFlow.Value.JobNames.Exists(x => x.Equals("All", StringComparison.CurrentCultureIgnoreCase)))
                {
                    return name;
                }
            }
            return "";
        }

        private static string FindDcSpec(string job)
        {
            if (TestProgram.IgxlWorkBk.DcSpecSheets.Count == 0)
            {
                return "";
            }

            var sheetNames = new List<string>();
            foreach (KeyValuePair<string, DcSpecSheet> dcSheet in TestProgram.IgxlWorkBk.DcSpecSheets)
            {
                string sheetName = dcSheet.Value.Name;
                string dcSpecJob = sheetName.Split('_').Count() > 2 ? sheetName.Split('_')[2] : "";
                if (dcSpecJob.Equals(job, StringComparison.OrdinalIgnoreCase))
                {
                    sheetNames.Add(sheetName);
                }
            }
            if (sheetNames.Any())
            {
                return string.Join(",", sheetNames);
            }

            return TestProgram.IgxlWorkBk.DcSpecSheets.FirstOrDefault().Value.Name;
        }
    }
}
