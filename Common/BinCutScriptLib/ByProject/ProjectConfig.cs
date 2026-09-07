using System;
using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Static;

using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.ByProject
{
    public class ProjectConfig
    {
        public virtual BinCutConfigXml GetBinCutConfig()
        {
            var config = new BinCutConfigXml();
            config.IdsNameList.Add(new Tuple<string, string>("IDS_VDD_FIXED", "VDD_FIXED_ALL"));
            config.IdsNameList.Add(new Tuple<string, string>("IDS_VDD_LOW", "VDD_LOW_ALL"));
            return config;
        }

        public virtual int GetHarvFailCount(Dictionary<string, string> harvResult)
        {
            List<string> harvesFlags = GetHarvFlags();
            return harvesFlags.Where(harvResult.ContainsKey).Count(flag => harvResult[flag] == "T");
        }

        public virtual List<string> GetHarvFlags()
        {
            return BinCutConfig.HarvFlags.Count != 0 ? BinCutConfig.HarvFlags.First().Value : null!;
        }

        public virtual Dictionary<string, string> HarvMappingTable(string pwrName)
        {
            return null!;
        }

        public virtual string GetIdsNameWithHarv(string powerName, Dictionary<string, string> harvesFlags)
        {
            return null!;
        }

        protected string GetIdsNameWithHarvBase(string powerName, Dictionary<string, string> harvesFlags, Func<string, bool>? powerNameFilter, Func<int, int, string> onPassIdsName)
        {
            if (powerNameFilter != null && !powerNameFilter(powerName))
            {
                return "";
            }

            IEnumerable<string> flags = HarvMappingTable(powerName).ToList().FindAll(x => x.Value == powerName).Select(y => y.Key);
            int passCount = 0;
            int coreCount = 0;
            string idsName = "";
            foreach (string key in flags)
            {
                if (!harvesFlags.ContainsKey(key))
                {
                    continue;
                }

                coreCount++;
                if (harvesFlags[key] == "F")
                {
                    passCount++;
                }
            }
            if (coreCount != 0 && coreCount - passCount <= 1) // 1 core fail or all core pass
            {
                idsName = onPassIdsName(coreCount, passCount);
            }

            return idsName;
        }
    }
}
