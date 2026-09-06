using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;


namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class PatternRow
    {
        #region Property
        public string SheetName { get; set; }
        public int RowNum { get; set; }
        public int PatternColumnNum { get; set; }
        public string TtrStr { get; set; }
        public string PartStr { get; set; }
        public string NoBinOutStr { get; set; }
        public string Enable { get; set; }
        public string SiteFlag { get; set; }
        public string FailFlag { get; set; }
        public string Description { get; set; }
        public PatternClass Pattern { get; set; }
        public ForceClass ForceCondition { get; set; }
        public string SpecifyTestName { get; set; }
        public string RfInterPose { get; set; }
        public string PostPatForceCondition { get; set; }
        public string RegisterAssignment { get; set; }
        public string MiscInfo { get; set; }
        public List<PatChildRow> PatChildRows { get; set; }
        public int DupIndex { get; set; }
        public bool IsRf { get; set; }
        #endregion

        public PatternRow(string patternName = "")
        {
            PatChildRows = new List<PatChildRow>();
            RowNum = 0;
            PatternColumnNum = 0;
            TtrStr = "";
            PartStr = "";
            NoBinOutStr = "";
            Enable = "";
            SiteFlag = "";
            FailFlag = "";
            Description = "";
            ForceCondition = new ForceClass();
            PostPatForceCondition = "";
            RegisterAssignment = "";
            MiscInfo = "";
            RfInterPose = "";
            DupIndex = 0;
            Pattern = new PatternClass(patternName);
        }

        public PatternRow(PatternRow other)
        {
            if (other == null)
            {
                return;
            }

            // Copy value types and strings
            SheetName = other.SheetName;
            RowNum = other.RowNum;
            PatternColumnNum = other.PatternColumnNum;
            TtrStr = other.TtrStr;
            PartStr = other.PartStr;
            NoBinOutStr = other.NoBinOutStr;
            Enable = other.Enable;
            SiteFlag = other.SiteFlag;
            FailFlag = other.FailFlag;
            Description = other.Description;
            SpecifyTestName = other.SpecifyTestName;
            RfInterPose = other.RfInterPose;
            PostPatForceCondition = other.PostPatForceCondition;
            RegisterAssignment = other.RegisterAssignment;
            MiscInfo = other.MiscInfo;
            DupIndex = other.DupIndex;
            IsRf = other.IsRf;

            Pattern = other.Pattern?.Copy();
            ForceCondition = other.ForceCondition?.Copy();

            PatChildRows = other.PatChildRows?.Select(row => row.Copy()).ToList() ?? new List<PatChildRow>();
        }

        public PatternRow Copy()
        {
            return new PatternRow(this);
        }

        public List<TestPlanRow> GetTestPlanRows()
        {
            List<TestPlanRow> testPlanRows = new List<TestPlanRow>();
            foreach (PatChildRow measRow in PatChildRows)
            {
                List<TestPlanRow> tpRows = ((PatSubChildRow)measRow).TpRows;
                testPlanRows.AddRange(tpRows);
            }
            return testPlanRows;
        }
    }
}
