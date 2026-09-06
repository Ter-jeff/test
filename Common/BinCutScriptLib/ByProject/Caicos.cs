using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.ByProject
{
    public class Caicos : ProjectConfig
    {
        public override BinCutConfigXml GetBinCutConfig()
        {
            var config = new BinCutConfigXml
            {
                IdsNameList =
                [
                    new("IDS_VDD_ECPU_25C", "VDD_ECPU"),
                    new("IDS_VDD_LPEM_25C", "VDD_LPEM"),
                    new("IDS_VDD_WARM_25C", "VDD_WARM"),
                    new("IDS_VDD_DCS_25C", "VDD_DCS"),
                    new("IDS_VDD_SOC_25C", "VDD_SOC"),
                    new("IDS_VDD_AON_25C", "VDD_AON"),
                    new("IDS_VDD_ALLI_25C", "VDD_ALLI"),
                    new("IDS_VDD_ECPU_SRAM_25C", "VDD_ECPU_SRAM"),
                    new("IDS_VDD_LPEM_SRAM_25C", "VDD_LPEM_SRAM"),
                    new("IDS_VDD_WARM_SRAM_25C", "VDD_WARM_SRAM"),
                    new("IDS_VDD_SOC_SRAM_25C", "VDD_SOC_SRAM"),
                    new("IDS_VDD_AON_SRAM_25C", "VDD_AON_SRAM"),
                ]
            };
            return config;
        }

    }
}
