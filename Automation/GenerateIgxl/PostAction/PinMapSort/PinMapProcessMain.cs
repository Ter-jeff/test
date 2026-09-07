using Automation.Static;

namespace Automation.GenerateIgxl.PostAction.PinMapSort
{
    public class PinMapProcessMain
    {
        public void WorkFlow()
        {
            var sortPin = new SortPinMap();
            if (TestProgram.IgxlWorkBk.PinMapPair.Value != null)
            {
                sortPin.Sort(TestProgram.IgxlWorkBk.PinMapPair.Value);
            }
        }
    }
}
