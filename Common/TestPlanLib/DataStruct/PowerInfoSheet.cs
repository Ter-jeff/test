using System.Collections.Generic;

using CommonReaderLib;

namespace TestPlanLib.DataStruct
{
    public class PowerInfoSheet : MySheet
    {
        public bool ExistEvs = false;
        public bool ExistConti = false;
        public bool ExistHip = false;

        public List<PowerInfoRow> Rows { get; set; } = [];
    }
}
