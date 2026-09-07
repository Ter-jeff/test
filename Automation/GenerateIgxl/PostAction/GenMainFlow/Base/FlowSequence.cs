using System.Text.RegularExpressions;

using IgxlLib.Enums;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow.Base
{
    public class FlowSequence
    {
        public string Job;
        public string SubFlowName;
        public string Enable;
        public EnumGroupInMainFlow GroupNameInMainFlow = EnumGroupInMainFlow.None;
        public string GroupSheetName;
        public int RowNum;

        public FlowSequence(string flow)
        {
            Job = "";
            if (flow.Contains(':'))
            {
                GroupSheetName = flow.Split(':')[0];
                SubFlowName = flow.Split(':')[1];
            }
            else
            {
                GroupSheetName = "";
                SubFlowName = flow;
            }
            const string regexPattern = @"\((?<enable>.+)\)";
            if (Regex.IsMatch(flow, regexPattern, RegexOptions.IgnoreCase))
            {
                string enable = Regex.Match(flow, regexPattern, RegexOptions.IgnoreCase).Groups["enable"].ToString();
                Enable = enable.Replace(" ", "_");
            }
        }
    }
}
