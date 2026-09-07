using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.ByProject
{
    public class Bora : ProjectConfig
    {
        public override BinCutConfigXml GetBinCutConfig()
        {
            var config = new BinCutConfigXml
            {
                IdsNameList =
                [
                    new("IDS_VDD_CPU_SRAM", "SRAM_CPU"),
                    new("IDS_VDD_GPU_SRAM", "SRAM_GPU"),
                    new("IDS_VDD_SOC_SRAM", "SRAM_SOC"),
                    new("IDS_VDD_ANE_SRAM", "SRAM_ANE"),
                ]
            };
            return config;
        }
    }
}
