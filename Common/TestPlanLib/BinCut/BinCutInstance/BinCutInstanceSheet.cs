using System.Collections.Generic;

using CommonReaderLib;

using TestPlanLib.EVS;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class BinCutInstanceSheet : MySheet
    {
        public List<BinCutInstanceRow> Rows = [];

        public int FlowNameColNumber = -1;
        public int InstanceColNumber = -1;
        public int OpcodeColNumber = -1;
        public int TestingStageColNumber = -1;
        public int EnableColNumber = -1;
        public int BinTableEnableColNumber = -1;
        public int PatternStartColNumber = -1;
        public int CommnetColNumber = -1;
        public int PatSetNameOrangeColNumber = -1;
        public int SubFlowColNumber = -1;
        public int EnableNewColNumber = -1;
        public int JobTestStageColNumber = -1;
        public int SiteVarColNumber = -1;
        public int FailFlagColNumber = -1;
        public int PassFlagColNumber = -1;
        public int BinOutStageColNumber = -1;
        public int VoltageColNumber = -1;
        public int LevelsColNumber = -1;
        public List<EvsConditionColIdx> EvsConditionColIdxDic = [];
        public int EvsVoltageColNumber = -1;
        public int EvsRampingColNumber = -1;
        public int EvsStressTimeColNumber = -1;
        public int EvsCoolingTimeColNumber = -1;
        public int EvsPwrPin1ColNumber = -1;
        public int EvsPwrPin2ColNumber = -1;
        public int EvsParallelSetting = -1;
        public int TypeColNumber = -1;
        public int DvsCoolingPatIndexColNumber = -1;
        public int DvsCoolingPatSetLoopNumberColNumber = -1;
        public int TimeSetColNumber = -1;
        public int ShiftSpeedColNumber = -1;
        public int PatternPinGroupColNumber = -1;
        public int PinGroupBinoutColNumber = -1;
        public int TempMonColNumber = -1;

        public int PerfModeColNumber = -1;
        public int PmOrderColNumber = -1;
        public int CallFlowColNumber = -1;

        public int MultiFstpEnableColNumber = -1;
        public int UserFunctionColNumber = -1;
        public int FunctionNameColNumber = -1;
        public int BurstColNumber = -1;

        public int SpecialSettingNumber = -1;
        public int OverlayNumber = -1;
        public int CharNumber = -1;

        #region Constructor
        public BinCutInstanceSheet(string sheetName)
        {
            SheetName = sheetName;
        }
        #endregion

        public bool HasPatternPinGroups()
        {
            return Rows.Exists(x => !string.IsNullOrEmpty(x.PatternPinGroup));
        }
    }
}
