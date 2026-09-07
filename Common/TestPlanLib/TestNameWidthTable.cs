using System.Collections.Generic;

namespace TestPlanLib
{
    public class TestNameWidthTable
    {
        public List<TestNameWidthRow> Rows = [];
    }

    public class TestNameWidthRow
    {
        public string Module = "";
        public string PtrTestNameWidt = "";
        public string PtrPinWidth = "";
        public string FtrTestNameWidt = "";
        public string FtrPatternWidth = "";

        public string Comment { get; internal set; } = "";
    }
}
