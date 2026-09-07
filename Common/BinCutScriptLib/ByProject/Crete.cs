using System.Collections.Generic;

namespace BinCutScriptLib.ByProject
{
    public class Crete : ProjectConfig
    {
        public override int GetHarvFailCount(Dictionary<string, string> harvResult)
        {
            int failCount = 0;
            List<string> harvesFlags = GetHarvFlags();
            foreach (string flag in harvesFlags)
            {
                if (!harvResult.TryGetValue(flag, out string? value))
                {
                    continue;
                }

                if (value == "T")
                {
                    failCount++;
                }
            }
            return failCount;
        }

        public override List<string> GetHarvFlags()
        {
            int[] gpuCoreNum = [6];
            int flagIdx = 0;
            var harvFlags = new List<string>();
            for (int j = flagIdx; j < gpuCoreNum[0]; j++)
            {
                harvFlags.Add("F_Gfx_Core" + flagIdx);
                flagIdx++;
            }
            return harvFlags;
        }

        public override Dictionary<string, string> HarvMappingTable(string pwrName)
        {
            var table = new Dictionary<string, string>();
            List<string> harvFlags = GetHarvFlags();
            foreach (string flag in harvFlags)
            {
                table.Add(flag, pwrName);
            }
            return table;
        }

        public override string GetIdsNameWithHarv(string powerName, Dictionary<string, string> harvesFlags)
        {
            return GetIdsNameWithHarvBase(
                powerName,
                harvesFlags,
                p => p == "VDD_GPU" || p == "VDD_SRAM_GPU",
                (coreCount, passCount) => "IDS_" + powerName + "_5"
            );
        }
    }
}
