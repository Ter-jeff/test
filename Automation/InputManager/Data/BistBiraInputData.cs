using System.Collections.Generic;

using Automation.GenerateIgxl.BistBira.Base;
using Automation.GenerateIgxl.BistBira.BistInputLib;
using Automation.Reader.ConfigFile.NamingRule.Base;

using ScghLib.Reader;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.Flow;

namespace Automation.InputManager.Data
{
    public class BistBiraInputData : InputDataBase
    {
        public BistNaming Naming { get; set; }
        public BinCutFlowTable EquationVoltage { get; set; }
        public MbistConfig Config { get; set; }
        public Dictionary<string, PatternData> PatternDic { get; set; }
        public PerformanceModeFilter PerformanceModeFilter { get; set; }
        public VoltageConverter VoltageConverter { get; set; }
        public List<BistProdFlowSheet> ProdFlowSheets { get; set; } = new List<BistProdFlowSheet>();
    }
}
