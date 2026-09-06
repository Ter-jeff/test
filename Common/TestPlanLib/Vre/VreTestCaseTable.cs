using System.Collections.Generic;
using System.Linq;

namespace TestPlanLib
{
    public class VreTestCaseTable
    {
        public List<VreTestCaseRow> Rows = new List<VreTestCaseRow>();
        public Dictionary<string, int> HeaderIndex;
        public string SheetName;
        public int GetMaxCaseId
        {
            get
            {
                return Rows.Where(x => x.CaseId != -1).Max(x => x.CaseId);
            }
        }
    }

    public class VreTestCaseRow
    {
        public int CaseId { get; set; } = -1;
        public string SubProgram { get; set; }
        public string ProcesureName { get; set; }
        public string InstanceName { get; set; }
        public List<string> Pattern { get; set; } = new List<string>();
        public string UserDef { get; set; }
        public string HardBin { get; set; }
        public string SorfdBin { get; set; }
        public string Comment { get; set; }
        public string LevelCheck { get; set; }

    }
}
