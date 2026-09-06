using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

namespace IgxlLib.Utility.Tester
{
    public static class TesterConfigManager
    {
        public static string GetToolTypeByChannelAssignment(Dictionary<string, TesterConfig> testerConfigDic, string channelAssignment, string sheetName)
        {
            if (channelAssignment != null)
            {
                TesterConfig? testerConfig = GetTesterConfigs(testerConfigDic, sheetName);
                if (testerConfig == null)
                {
                    return "";
                }

                return testerConfig.GetToolType(channelAssignment.Split('.')[0]);
            }
            return "";
        }

        private static TesterConfig? GetTesterConfigs(Dictionary<string, TesterConfig> testerConfigDic, string sheetName)
        {
            string name = sheetName.ToUpper();
            if (name.Contains("_CP") && testerConfigDic.TryGetValue("CP", out TesterConfig? cpConfig))
            {
                return cpConfig;
            }

            if (name.Contains("_FT") && testerConfigDic.TryGetValue("FT", out TesterConfig? ftConfig))
            {
                return ftConfig;
            }

            if (testerConfigDic.TryGetValue("All", out TesterConfig? allConfig))
            {
                return allConfig;
            }

            if (testerConfigDic.TryGetValue("CP", out TesterConfig? defaultCpConfig))
            {
                return defaultCpConfig;
            }

            return null;
        }

        public static string GetToolTypeByChannelAssignments(Dictionary<string, TesterConfig> testerConfigDic, List<string> channelAssignments, string sheetName)
        {
            var toolType = channelAssignments.Where(x => !x.EqualsIgnoreCase("SITE0")).Select(x => GetToolTypeByChannelAssignment(testerConfigDic, x, sheetName)).ToList();
            if (toolType.Distinct().Count() == 1 && toolType.Count > 0)
            {
                return toolType.First();
            }

            return "";
        }
    }
}
