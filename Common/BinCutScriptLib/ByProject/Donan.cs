using TestPlanLib.BinCut.BinCutConfig;

namespace BinCutScriptLib.ByProject
{
    public class Donan : ProjectConfig
    {
        public override BinCutConfigXml GetBinCutConfig()
        {
            var config = new BinCutConfigXml
            {
                IdsNameList =
                [
                    new("IDS_VDD_LOW", "VDD_LOW_ALL"),
                    new("IDS_VDD_FIXED", "VDD_FIXED_ALL"),
                ]
            };
            return config;
        }
    }
}
