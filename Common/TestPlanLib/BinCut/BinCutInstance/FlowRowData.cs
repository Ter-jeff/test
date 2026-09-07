namespace TestPlanLib.BinCut.BinCutInstance
{
    public class FlowRowData
    {
        public string FlowNameOri = "";
        public string Instance = "";
        public string Opcode = "";
        public string EnableAndDevice = "";
        public string BintableEnableWd = "";
        public string SubFlow = "";
        public string EnableFlow = "";
        public string JobTestStage = "";
        public string SiteVar = "";
        public string FailFlag = "";
        public string PassFlag = "";
        public string BinOutStage = "";
        public string TempMon = "";
        public string DCcategory = "";
        public string Levels = "";
        public string HarvPinGrpEnable = "";
        public string HarvestBinningFlag = "";
        public string Burst = "";
        public string TimeSet = "";
        public string ShiftSpeed = "";
        public string PatSetNameOrange = "";

        public FlowRowData Copy()
        {
            return new FlowRowData
            {
                FlowNameOri = FlowNameOri,
                Instance = Instance,
                Opcode = Opcode,
                EnableAndDevice = EnableAndDevice,
                BintableEnableWd = BintableEnableWd,
                SubFlow = SubFlow,
                EnableFlow = EnableFlow,
                JobTestStage = JobTestStage,
                SiteVar = SiteVar,
                FailFlag = FailFlag,
                PassFlag = PassFlag,
                BinOutStage = BinOutStage,
                TempMon = TempMon,
                DCcategory = DCcategory,
                Levels = Levels,
                HarvPinGrpEnable = HarvPinGrpEnable,
                HarvestBinningFlag = HarvestBinningFlag,
                Burst = Burst,
                TimeSet = TimeSet,
                ShiftSpeed = ShiftSpeed,
                PatSetNameOrange = PatSetNameOrange
            };
        }
    }
}
