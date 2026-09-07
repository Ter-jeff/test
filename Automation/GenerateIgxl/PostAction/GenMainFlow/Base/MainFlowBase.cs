using System.Collections.Generic;
using System.Linq;

using Automation.Static;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow.Base
{
    public class MainFlowBase
    {
        public string MainFlowName { get; set; } = "";
        public string JobName { get; set; } = "";
        public string SheetName { get; set; } = "";
        public int JobColIndex { get; set; }
        private string _igxlJobName = "";
        public string IgxlJobName
        {
            get
            {
                if (string.IsNullOrEmpty(_igxlJobName))
                {
                    _igxlJobName = GetIgxlJobName(JobName);
                }
                return _igxlJobName;
            }
        }
        public string GetIgxlJobName(string jobName)
        {
            string realJobName = jobName;
            TestPlanStatic.TestProgramDefSheet?.JobMappingDic.TryGetValue(jobName, out realJobName);
            if (string.IsNullOrEmpty(realJobName))
            {
                realJobName = jobName;
            }
            return realJobName;
        }
        public List<string> GetAllGroups
        {
            get
            {
                List<string> allGroups = SequencesNew.Where(x => !string.IsNullOrEmpty(x.Group)).Select(x => x.Group).ToList();
                return allGroups;
            }
        }

        public List<FlowSequence> Sequences { get; set; } = new List<FlowSequence>();
        public List<FlowSequenceNew> SequencesNew { get; set; } = new List<FlowSequenceNew>();
        public List<string> AvalibleParts = new List<string>();
    }
}
