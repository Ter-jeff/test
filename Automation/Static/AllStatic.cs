using Automation.Static.Result;

using CommonLib.ErrorReport;

using TestPlanLib.Static;

namespace Automation.Static
{
    public class AllStatic
    {
        public static void Clear()
        {
            ErrorReportManager.ClearErrors();
            BasicResult.Clear();
            EfuseResult.Clear();

            TestPlanStatic.Clear();
            ScghStatic.Clear();
            SettingStatic.Clear();
            TestProgram.Clear();
            SettingStatic.Clear();
            LocalSpecs.Clear();
            FolderStructure.Clear();
            EpWorkbook.Clear();
            BlockStatus.Clear();

            LocalSpecs.Clear();
            TestProgram.Clear();
            EpWorkbook.Clear();
            VbtFunctionLibShared.Clear();
            HardIpStatic.Clear();
            EfuseStatic.Clear();
        }
    }
}
