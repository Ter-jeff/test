using System.Collections.Generic;

using TestPlanLib.BinCut.BinCutInstance;

namespace Automation.InputManager.Data
{
    public class EvsInputData : InputDataBase
    {
        public List<BinCutInstanceSheet> EvsInstanceSheets { get; set; }
    }
}
