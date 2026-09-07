using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;
using Cautogen.common.IgxlProgramLib.IgxlProgramParser;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class UpdateJobListSheet
    {
        public static void WorkFlow(IgxlProgram mTestProgram)
        {
            GeneralFunc.WriteMessage("Update JobList sheet... ");

            JobListSheet jobSheet = mTestProgram.JoblistSheet;
            string czFolder = Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder);
            // create job folder
            string jobFileName = Path.Combine(LocalSpecs.OutputFolder, ConstData.CommonFolder, "Jobs", jobSheet.Name + ".txt");
            string jobFolder = Path.GetDirectoryName(jobFileName);
            if (jobFolder != null)
            {
                if (!Directory.Exists(jobFolder))
                {
                    Directory.CreateDirectory(jobFolder);
                }
            }

            foreach (JobRow job in jobSheet.Rows)
            {
                job.PinMap = string.Join(",", job.PinMap);
                job.FlowTable = job.FlowTable;
                job.AcSpecs = string.Join(",", job.AcSpecs);
                job.DcSpecs = string.Join(",", job.DcSpecs);
                job.PatternGroups = string.Join(",", job.PatternGroups);
                job.TestProcedures = string.Join(",", job.TestProcedures);
                job.MixedSignalTiming = string.Join(",", job.MixedSignalTiming);
                job.WaveDefinitions = string.Join(",", job.WaveDefinitions);
                job.PSets = string.Join(",", job.PSets);
                job.Signals = string.Join(",", job.Signals);
                job.PortMap = string.Join(",", job.PortMap);
                job.FractionalBus = string.Join(",", job.FractionalBus);
                job.ConcurrentSequence = string.Join(",", job.ConcurrentSequence);
                job.TestInstances = job.TestInstances;
                job.BinTable = job.BinTable;
                job.PatternSets = string.Join(",", GetPatternSets(job));
                job.Characterization = GetCharSheets(job.Characterization.Split(','), czFolder);
            }
            jobSheet.Write(jobFileName);
            LocalSpecs.GenSheets.Add(jobSheet);
        }

        private static List<string> GetPatternSets(JobRow job)
        {
            var patternSets = new List<string>();
            foreach (string item in job.PatternSets.Split(',').ToList())
            {
                if (!item.Equals("PatSets_All", StringComparison.CurrentCultureIgnoreCase))
                {
                    patternSets.Add(item);
                }
            }

            if (LocalSpecs.ProgramUpdateOnly)
            {
                patternSets.Add("PatSets_All");
            }
            else
            {
                patternSets.Add("PatSets_All_CZ");
            }

            if (LocalSpecs.InputParam.UseNewTChar)
            {
                patternSets.Add("PatSets_CZ");
            }

            return patternSets;
        }

        private static string GetNewInstanceList(List<string> instanceList, string czFolder)
        {
            if (LocalSpecs.AllModuleSheets == null && LocalSpecs.AllCommonSheets == null)
            {
                return string.Join(",", instanceList);
            }

            var newList = new List<string>();
            newList = instanceList;
            //newList.AddRange(instanceList.Where(a => LocalSpecs.AllModuleSheets.Contains(a)));  //Remove un-used instance sheets
            //newList.AddRange(instanceList.Where(a => LocalSpecs.AllCommonSheets.Contains(a)));  //Remove un-used instance sheets ==== TestInst_Common
            var files = Directory.GetFiles(czFolder).Select(Path.GetFileNameWithoutExtension)
                    .Where(a => Regex.IsMatch(a, "^TestInst_", RegexOptions.IgnoreCase))
                    .ToList();
            newList.AddRange(files);//Add all CZ instance sheets
            return string.Join(",", newList);
        }

        private static string GetBinTableSheets(IEnumerable<string> binTableList)
        {
            if (LocalSpecs.AllModuleSheets == null && LocalSpecs.AllCommonSheets == null)
            {
                return string.Join(",", binTableList);
            }

            var newBinTableList = binTableList.Where(LocalSpecs.AllCommonSheets.Contains).ToList();//Remove un-used bintable sheets
            if (newBinTableList == null || newBinTableList.Count == 0)
            {
                return "";
            }

            return string.Join(",", newBinTableList);
        }

        private static string GetCharSheets(IEnumerable<string> charList, string czFolder)
        {
            if (LocalSpecs.AllModuleSheets == null && LocalSpecs.AllCommonSheets == null)
            {
                return string.Join(",", charList);
            }

            var newList = new List<string>();
            newList.AddRange(charList.Where(LocalSpecs.AllCommonSheets.Contains));//Remove un-used cz sheets
            var files = Directory.GetFiles(czFolder).Select(Path.GetFileNameWithoutExtension)
                    .Where(a => Regex.IsMatch(a, "^DevChar_", RegexOptions.IgnoreCase))
                    .ToList();
            newList.AddRange(files); //Add all CZ instance sheets
            return string.Join(",", newList);
        }
    }
}
