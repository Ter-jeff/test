using System.Collections.Generic;

namespace TestPlanLib.BinNumber
{
    public class SofBinTable
    {
        public class SofBintable(string sofSheetName, List<string> bintables)
        {
            public string SofSheetName = sofSheetName;
            public List<string> Bintables = bintables ?? [];
        }
    }
}
