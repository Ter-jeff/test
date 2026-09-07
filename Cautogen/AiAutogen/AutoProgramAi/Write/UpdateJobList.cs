using System.Linq;

using IgxlLib.IgxlSheets;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateJobList
    {
        public JobListSheet Work(JobListSheet jobListSheet, PatSetSheet patSetsAllCz,
            PatSetSheet patSetCz, InstanceSheet instanceSheet, CharSheet charSheet)
        {
            foreach (var row in jobListSheet.Rows)
            {
                var testInstances = row.TestInstances.Split(',').ToList();
                if (instanceSheet.Rows.Any())
                {
                    testInstances.Add(instanceSheet.Name);
                }
                row.TestInstances = string.Join(",", testInstances.Where(x => !string.IsNullOrEmpty(x)));
                var patternSets = row.PatternSets.Split(',').ToList();
                if (patSetsAllCz.Rows.Any())
                {
                    patternSets.Add(patSetsAllCz.Name);
                }
                if (patSetCz.Rows.Any())
                {
                    patternSets.Add(patSetCz.Name);
                }
                row.PatternSets = string.Join(",", patternSets.Where(x => !string.IsNullOrEmpty(x)));
                var characterizations = row.Characterization.Split(',').ToList();
                if (charSheet.Rows.Any())
                {
                    characterizations.Add(charSheet.Name);
                }
                row.Characterization = string.Join(",", characterizations.Where(x => !string.IsNullOrEmpty(x)));
            }

            return jobListSheet;
        }
    }
}
