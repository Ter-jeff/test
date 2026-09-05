namespace TestPlanLib.Utility
{
    public class SubBinCutSheetInfo
    {
        public int StartColumnNum { set; get; }
        public int SubFlowEndColumnNum { set; get; }
        public int EndColNum { set; get; }
        public string JobName { set; get; } = "";
        public int JobColumnNum { set; get; }

        public SubBinCutSheetInfo()
        {
        }

        public SubBinCutSheetInfo(SubBinCutSheetInfo subBinCutSheetInfo)
        {
            if (subBinCutSheetInfo == null)
            {
                return;
            }

            StartColumnNum = subBinCutSheetInfo.StartColumnNum;
            SubFlowEndColumnNum = subBinCutSheetInfo.SubFlowEndColumnNum;
            EndColNum = subBinCutSheetInfo.EndColNum;
            JobName = subBinCutSheetInfo.JobName;
            JobColumnNum = subBinCutSheetInfo.JobColumnNum;
        }

        public SubBinCutSheetInfo Copy()
        {
            return new SubBinCutSheetInfo(this);
        }
    }
}
