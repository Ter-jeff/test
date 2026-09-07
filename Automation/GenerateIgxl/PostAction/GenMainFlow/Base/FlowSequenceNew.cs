using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Automation.GenerateIgxl.PostAction.GenMainFlow.Base
{
    public class FlowSequenceNew
    {
        private string _option = "";
        public string Job;
        public string OriSheetName;
        public string Part;
        public bool Enable = false;
        public int RowNum { get; set; }
        public string Group;
        public bool Found = false;

        public string Source { get; set; }
        public string Module { get; set; }
        public string SheetName { get; set; }
        public string SubFlowName { get; set; }
        public string SubProgramFlowName { get; set; }
        public string EnableWord { get; set; }
        public string BintableEnableWord { get; set; }
        public string SiteFlagPerSite { get; set; }
        public string FailFlag { get; set; }
        public string Option
        {
            get { return _option; }
            set
            {
                _option = value;
                Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                string regexPattern = @"(?<key>[^=:;,\s]+)\s*[=:]\s*(?<value>(?:""[^""]*"")|[^;,\s]+)|(?<key>\w+)";
                MatchCollection matches = Regex.Matches(_option, regexPattern);
                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        string inputKey = match.Groups["key"].Value.Trim();
                        string inputValue = match.Groups["value"].Success ? match.Groups["value"].Value.Trim().Trim('"').Trim() : "";
                        if (!match.Groups["value"].Success)
                        {
                            inputValue = inputKey;
                        }
                        if (result.ContainsKey(inputKey))
                        {
                            result[inputKey] += ";" + inputValue;
                        }
                        else
                        {
                            result.Add(inputKey, inputValue);
                        }
                    }
                }
                OptionDict = result;
            }
        }
        public Dictionary<string, string> OptionDict { get; private set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> JobDic { get; set; } = new Dictionary<string, string>();
        public string Comment { get; set; }
        public bool IsEvsDeferredBinout
        {
            get
            {
                if (!string.IsNullOrEmpty(Option) && Option.IndexOf("EvsDeferredBinout", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    return true;
                }

                return false;
            }
        }

        internal string GetJobs()
        {
            if (JobDic.Values.All(x => x.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase)))
            {
                return "";
            }
            return string.Join(",", GetJobList());
        }

        internal List<string> GetJobList()
        {
            List<string> flags = JobDic.Where(x => x.Value.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase)).Select(x => x.Key).ToList();
            return flags;
        }
    }
}
