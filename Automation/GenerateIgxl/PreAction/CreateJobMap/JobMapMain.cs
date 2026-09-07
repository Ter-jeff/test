using Automation.Static;

using TestPlanLib.Settings;

namespace Automation.GenerateIgxl.PreAction.CreateJobMap
{
    public class JobMapMain
    {
        public void WorkFlow()
        {
            JobMapSheet jobMapSheet = SettingStatic.JobMapSheet;
            LocalSpecs.JobMap = jobMapSheet.JobMapDictionary;
            LocalSpecs.JobTemperatureMap = jobMapSheet.JobTempMap;
        }
    }
}
