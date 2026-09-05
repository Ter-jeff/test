using System.Collections.Generic;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow
{
    public class MainFlowData
    {
        public string FullSheetName;
        public List<string> JobNames { get; set; } //List for one sheet can used for two job in main flow
    }
}
