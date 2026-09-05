using System;
using System.Collections.Generic;
using System.Linq;

using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.ByProject
{
    public class Ellis : ProjectConfig
    {
        public override int GetHarvFailCount(Dictionary<string, string> harvResult)
        {
            List<string> harvesFlags = GetHarvFlags();
            return harvesFlags.Where(harvResult.ContainsKey).Count(flag => harvResult[flag] == "T");
        }

        public override List<string> GetHarvFlags()
        {
            int[] gpuCoreNum = [6];
            int flagIdx = 0;
            var harvFlags = new List<string>();
            for (int j = flagIdx; j < gpuCoreNum[0]; j++)
            {
                harvFlags.Add("F_Gfx_HARV" + flagIdx);
                harvFlags.Add("F_Gfx_CORE" + flagIdx);
                flagIdx++;
            }
            return harvFlags;
        }

        public override Dictionary<string, string> HarvMappingTable(string pwrName)
        {
            List<string> harvFlags = GetHarvFlags();
            return harvFlags.ToDictionary(flag => flag, flag => pwrName);
        }

        public override string GetIdsNameWithHarv(string powerName, Dictionary<string, string> harvesFlags)
        {
            return GetIdsNameWithHarvBase(powerName, harvesFlags, null, (coreCount, passCount) => "IDS_VDD_SRAM_GPU_" + passCount);
        }

        public override BinCutConfigXml GetBinCutConfig()
        {
            var config = new BinCutConfigXml
            {
                IdsNameList =
                [
                    new("IDS_VDD_GPU_BINCHECK", "VDD_GPU"),
                    new("ids_vdd_sram_gpu_4", "VDD_SRAM_GPU"),
                    new("IDS_VDD_SOC_SRAM", "SRAM_SOC"),
                    new("IDS_VDD_ANE_SRAM", "SRAM_ANE"),
                ]
            };
            return config;
        }
    }
}
