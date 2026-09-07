using System.Collections.Generic;
using System.Data;

using Automation.GenerateIgxl.HardIp.InputObject;

namespace RfLib.InstrumentSetup
{
    public class InstrumentSetupRow
    {
        public string Pattern { get; set; } = "";
        public string Pins { get; set; } = "";
        public string MeasSeqType { get; set; } = "";
        public DataTable InstrumentDataTable { get; set; } = new();
        public MeasPin Pin { get; set; } = new();
        public int RowIndex { get; set; }
        public int SeqIndex { get; set; }
        public int SubsetIndex { get; set; }
        public string SetupName { get; set; } = "";
        public string SubSetting { get; set; } = "";
        public List<string> InstrumentType { get; set; } = [];

        public object Clone()
        {
            return MemberwiseClone();
        }
    }
}
