using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.ByProject
{
    public class Rhodes : ProjectConfig
    {
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
                    new("IDS_VDD_AVEMSR_25C", "VDD_AVEMSR"),
                    new("IDS_VDD_DISP_25C", "VDD_DISP"),
                    new("IDS_VDD_DISP2_25C", "VDD_DISP2"),
                    new("IDS_VDD_ANE_25C", "VDD_ANE"),
                    new("IDS_VDD_BUMPER_25C", "VDD_BUMPER"),
                    new("IDS_VDD_GPU0_BINCHECK", "VDD_GPU0"),
                    new("IDS_VDD_GPU1_BINCHECK", "VDD_GPU1"),
                    new("IDS_VDD_GPU2_BINCHECK", "VDD_GPU2"),
                    new("IDS_VDD_GPU3_BINCHECK", "VDD_GPU3"),
                    new("IDS_VDD_ECPU_25C", "VDD_ECPU"),
                    new("IDS_VDD_PCPU0_BINCHECK", "VDD_PCPU0"),
                    new("IDS_VDD_PCPU1_BINCHECK", "VDD_PCPU1"),
                    new("IDS_VDD_PCPU_SRAM0_25C", "VDD_PCPU_SRAM0"),
                    new("IDS_VDD_PCPU_SRAM1_25C", "VDD_PCPU_SRAM1"),
                    new("IDS_VDD_ECPU_SRAM_25C", "VDD_ECPU_SRAM"),
                    new("IDS_VDD_ANE_SRAM_25C", "VDD_ANE_SRAM"),
                    new("IDS_VDD_GPU_SRAM0_25C", "VDD_GPU_SRAM0"),
                    new("IDS_VDD_GPU_SRAM1_25C", "VDD_GPU_SRAM1"),
                    new("IDS_VDD_GPU_SRAM2_25C", "VDD_GPU_SRAM2"),
                    new("IDS_VDD_GPU_SRAM3_25C", "VDD_GPU_SRAM3"),
                    new("IDS_VDD_SRAM_25C", "VDD_SRAM"),
                    new("IDS_VDD_FIXED_25C", "VDD_FIXED"),
                    new("IDS_VDD_LOW_25C", "VDD_LOW"),
                    new("IDS_VDD_CIO_25C", "VDD_CIO"),
                ]
            };
            return config;
        }
    }
}
