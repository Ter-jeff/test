using System.Collections.Generic;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.InputManager.Data
{
    public class HtolInputData : InputDataBase
    {
        public List<BinCutInstanceSheet> BinCutInstanceSheets { get; set; } = new List<BinCutInstanceSheet>();
    }
}
