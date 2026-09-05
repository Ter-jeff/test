using System.Collections.Generic;

namespace TestPlanLib.EVS
{
    public class EvsCondition(string jobStage, string evsNum)
    {
        public string JobStage = jobStage.Trim();
        public string EvsNum = evsNum.Trim();
        public string NumConds = "";
        public string Pulses = "";
        public string Time = "";
        public string Voltage1 = "";
        public string Voltage2 = "";
        public string Cooling = "";
        public string CoolingAfter = "";
        public string TotalPwr = "";
        public string AlarmFlag = "";
        public string RisingDelayTimeSec = "";
        public Dictionary<string, string> UserFunction = [];

        public EvsCondition(EvsCondition evsCondition) : this(evsCondition?.JobStage ?? "", evsCondition?.EvsNum ?? "")
        {
            if (evsCondition == null)
            {
                return;
            }

            NumConds = evsCondition.NumConds;
            Pulses = evsCondition.Pulses;
            Time = evsCondition.Time;
            Voltage1 = evsCondition.Voltage1;
            Voltage2 = evsCondition.Voltage2;
            Cooling = evsCondition.Cooling;
            CoolingAfter = evsCondition.CoolingAfter;
            TotalPwr = evsCondition.TotalPwr;
            AlarmFlag = evsCondition.AlarmFlag;
            RisingDelayTimeSec = evsCondition.RisingDelayTimeSec;
            UserFunction = new Dictionary<string, string>(evsCondition.UserFunction);
        }

        public EvsCondition Copy()
        {
            return new EvsCondition(this);
        }
    }
}
