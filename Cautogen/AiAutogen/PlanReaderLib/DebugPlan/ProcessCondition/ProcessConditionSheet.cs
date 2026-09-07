using CommonReaderLib;

namespace DebugPlanReaderLib.DebugPlan
{
    public class ProcessConditionSheet : MySheet
    {
        public ProcessConditionSheet(string sheetName)
        {
            SheetName = sheetName;
        }

        public string EfuseEnableWord { get; set; }
        public string Tester { get; set; }
    }
}
