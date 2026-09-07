using System.Collections.Generic;

namespace IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet
{
    public class MultiTimeSetSheets
    {
        public List<ComTimeSetBasicSheet> TimeSetBasicSheetsList { get; } = [];

        public void AddTimeSetSheet(ComTimeSetBasicSheet comTimeSetBasicSheet)
        {
            TimeSetBasicSheetsList.Add(comTimeSetBasicSheet);
        }
    }
}
