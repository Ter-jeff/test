using Automation.Reader.ConfigFile.NamingRule.Base;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.CPM;
using TestPlanLib.Scan;

namespace Automation.InputManager.Data
{
    public class ScanInputData : InputDataBase
    {
        public ScanConfig ScanConfig { get; set; }
        public ClockMeasSheet ClockMeasSheet { get; set; }
        public EfuseCpmSheet EfuseCpmSheet { get; set; }
        public BinCutInstanceSheet TurboModeInstanceSheet { get; set; }
    }
}
