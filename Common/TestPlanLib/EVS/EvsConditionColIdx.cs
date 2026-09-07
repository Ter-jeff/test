namespace TestPlanLib.EVS
{
    public class EvsConditionColIdx(string jobStage, string evsNum)
    {
        public string JobStage = jobStage.Trim();
        public string EvsNum = evsNum.Trim();
        public int NumCondsIdx = -1;
        public int PulsesIdx = -1;
        public int TimeIdx = -1;
        public int Voltage1Idx = -1;
        public int Voltage2Idx = -1;
        public int CoolingIdx = -1;
        public int CoolingAfterIdx = -1;
        public int TotalPwrIdx = -1;
        public int AlarmFlagIdx = -1;
        public int RisingDelayTimeSecIdx = -1;
        public int RampIdx = -1;
        public int VtrigIdx = -1;
        public int IVdx = -1;
    }
}
