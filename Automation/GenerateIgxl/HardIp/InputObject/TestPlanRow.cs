using System.Collections.Generic;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class TestPlanRow
    {
        public int RowNum { get; set; }
        public string Description { get; set; }
        public string ForceCondition { get; set; }
        public string RegisterAssignment { get; set; }
        public string MiscInfo { get; set; }
        public string Meas { get; set; }
        public int SequenceIndex { get; set; }
        public List<MeasLimit> Limits { get; set; } = new List<MeasLimit>();
        public int MergeRowNumForMeas { get; set; }
        public string TestName { get; set; } = "";
        public string InterposeFunc { get; set; } = "";
        public string RfInterpose { get; set; }
        public string RfIntrumentSetup { get; set; } = "";
    }
}
