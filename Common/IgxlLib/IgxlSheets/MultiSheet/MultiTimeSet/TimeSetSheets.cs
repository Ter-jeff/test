using System.Collections.Generic;

namespace IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet
{
    public class TimeSetSheets : List<ComTimeSetBasicSheet>
    {
        public void AddTimeSetSheet(ComTimeSetBasicSheet comTimeSetBasicSheet)
        {
            Add(comTimeSetBasicSheet);
        }
    }
}
