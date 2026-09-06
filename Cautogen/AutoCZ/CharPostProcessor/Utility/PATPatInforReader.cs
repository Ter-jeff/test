using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility
{
    public class PatPatInforReader
    {
        public string VmVectorReader(List<string> list)
        {
            foreach (string line in list)
            {
                string context = line.Trim('\r').Trim();

                if (context == "")
                {
                    continue;
                }

                if (Regex.IsMatch(context, "Module names:", RegexOptions.IgnoreCase))
                {
                    var moduleNameList = context.Split(':')[1].Split(',').Select(x => x.Trim()).Where(x => !string.IsNullOrEmpty(x)).ToList();
                    return string.Join(",", moduleNameList);
                }
            }
            return "";
        }
    }
}
