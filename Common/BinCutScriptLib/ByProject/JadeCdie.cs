using System;
using System.Collections.Generic;
using System.Linq;

using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.ByProject
{
    public class JadeCdie : ProjectConfig
    {
        public override int GetHarvFailCount(Dictionary<string, string> harvResult)
        {
            List<string> harvesFlags = GetHarvFlags();
            return harvesFlags.Where(harvResult.ContainsKey).Count(flag => harvResult[flag] == "T");
        }

        public override List<string> GetHarvFlags()
        {
            return [.. HarvMappingTable().Select(x => x.Key)];
        }

        public override Dictionary<string, string> HarvMappingTable(string pwrName)
        {
            return HarvMappingTable();
        }

        public static Dictionary<string, string> HarvMappingTable()
        {
            var table = new Dictionary<string, string>();
            int[] gpuCoreNum = [8, 8, 8, 8];
            int[] ecpuCoreNum = [4, 4];
            int[] pcpuCoreNum = [4, 4];
            int flagIdx = 0;
            for (int i = 0; i < gpuCoreNum.Length; i++)
            {
                if (gpuCoreNum[i] < 0)
                {
                    continue;
                }

                for (int j = 0; j < gpuCoreNum[i]; j++)
                {
                    table.Add("F_Gfx_HARV" + flagIdx, "VDD_GPU_SRAM" + i);
                    table.Add("F_Gfx_CORE" + flagIdx, "VDD_GPU_SRAM" + i);
                    flagIdx++;
                }
            }
            flagIdx = 0;
            for (int i = 0; i < ecpuCoreNum.Length; i++)
            {
                if (ecpuCoreNum[i] < 0)
                {
                    continue;
                }

                for (int j = 0; j < ecpuCoreNum[i]; j++)
                {
                    table.Add("F_ECPU_HARV" + flagIdx, "VDD_ECPU_SRAM" + i);
                    table.Add("F_ECPU_CORE" + flagIdx, "VDD_ECPU_SRAM" + i);
                    flagIdx++;
                }
            }
            flagIdx = 0;
            for (int i = 0; i < pcpuCoreNum.Length; i++)
            {
                if (pcpuCoreNum[i] < 0)
                {
                    continue;
                }

                for (int j = 0; j < pcpuCoreNum[i]; j++)
                {
                    table.Add("F_PCPU_HARV" + flagIdx, "VDD_PCPU_SRAM" + i);
                    table.Add("F_PCPU_CORE" + flagIdx, "VDD_PCPU_SRAM" + i);
                    flagIdx++;
                }
            }
            return table;
        }

        public override string GetIdsNameWithHarv(string powerName, Dictionary<string, string> harvesFlags)
        {
            return GetIdsNameWithHarvBase(
                powerName,
                harvesFlags,
                null,
                (coreCount, passCount) => "IDS_" + powerName + "_25C_" + passCount
            );
        }

        public override BinCutConfigXml GetBinCutConfig()
        {
            var config = new BinCutConfigXml
            {
                IdsNameList =
                [
                    new("IDS_VDD_FABRIC_25C", "VDD_FABRIC"),
                    new("IDS_VDD_AFR_25C", "VDD_AFR"),
                    new("IDS_VDD_SOC_25C", "VDD_SOC"),
                    new("IDS_VDD_DCS_25C", "VDD_DCS"),
                    new("IDS_VDD_DISP_25C", "VDD_DISP"),
                    new("IDS_VDD_DISP2_25C", "VDD_DISP2"),
                    new("IDS_VDD_GPU0_BINCHECK", "VDD_GPU0"),
                    new("IDS_VDD_GPU1_BINCHECK", "VDD_GPU1"),
                    new("IDS_VDD_GPU2_BINCHECK", "VDD_GPU2"),
                    new("IDS_VDD_GPU3_BINCHECK", "VDD_GPU3"),
                    new("IDS_VDD_PCPU0_BINCHECK", "VDD_PCPU0"),
                    new("IDS_VDD_PCPU1_BINCHECK", "VDD_PCPU1"),
                    new("IDS_VDD_ANE0_25C", "VDD_ANE0"),
                    new("IDS_VDD_ANE1_25C", "VDD_ANE1"),
                    new("IDS_VDD_AVEMSR_25C", "VDD_AVEMSR"),
                    new("IDS_VDD_GPU_BMPR_25C", "VDD_GPU_BMPR"),
                    new("IDS_VDD_PCPU_SRAM0_25C_4", "VDD_PCPU_SRAM0"),
                    new("IDS_VDD_PCPU_SRAM1_25C_4", "VDD_PCPU_SRAM1"),
                    new("IDS_VDD_ECPU_SRAM_25C", "VDD_ECPU_SRAM"),
                    new("IDS_VDD_ANE_SRAM0_25C", "VDD_ANE_SRAM0"),
                    new("IDS_VDD_ANE_SRAM1_25C", "VDD_ANE_SRAM1"),
                    new("IDS_VDD_GPU_SRAM0_25C_8", "VDD_GPU_SRAM0"),
                    new("IDS_VDD_GPU_SRAM1_25C_8", "VDD_GPU_SRAM1"),
                    new("IDS_VDD_GPU_SRAM2_25C_8", "VDD_GPU_SRAM2"),
                    new("IDS_VDD_GPU_SRAM3_25C_8", "VDD_GPU_SRAM3"),
                    new("IDS_VDD_SRAM_25C", "VDD_SRAM"),
                    new("IDS_VDD_FIXED_25C", "VDD_FIXED"),
                    new("IDS_VDD_LOW_25C", "VDD_LOW"),
                    new("IDS_VDD_CIO_25", "VDD_CIO"),
                ]
            };
            return config;
        }
    }
}
