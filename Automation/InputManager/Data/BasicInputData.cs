using System.Collections.Generic;

using Automation.Reader;

using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using TestPlanLib.DataStruct;

namespace Automation.InputManager.Data
{
    public class BasicInputData : InputDataBase
    {
        public IoContiSheet IoContiSheet { get; set; } = new IoContiSheet();
        public IfoldPowerTableSheet IfoldPowerTableSheet { get; set; } = new IfoldPowerTableSheet("");
        public PinMapSheet PinMapSheet { get; set; } = new PinMapSheet("");
        public IoLevelsSheet IoLevelsSheet { get; set; } = new IoLevelsSheet("");
        public List<DcTestContiSheet> DcTestContiSheets { get; set; } = new List<DcTestContiSheet>();
        public TimeSetSheets MultiTimeSetSheets { get; set; } = new TimeSetSheets();
    }
}
