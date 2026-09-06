namespace Automation.Static.Result
{
    public static class BasicResult
    {
        public static bool CheckIoContinuity { get; set; }
        public static bool CheckDcTestContinuity { get; set; }
        public static bool CheckPowerInfo { get; set; }
        public static bool CheckIoInfo { get; set; }
        public static bool CheckCsvPattenList { get; set; }
        public static bool CheckDirTimeSet { get; set; }
        public static bool HasTimeSet { get; set; }
        public static bool HasDirPatSet { get; set; }

        public static bool Ac = true;
        public static bool Level = true;
        public static bool PatternList => CheckCsvPattenList;
        public static bool PattenSet => CheckCsvPattenList && HasDirPatSet;
        public static bool TimeSet => CheckCsvPattenList && CheckDirTimeSet;
        public static bool Continuity => CheckDcTestContinuity;

        public static void Clear()
        {
            CheckIoContinuity = true;
            CheckDcTestContinuity = true;
            CheckPowerInfo = true;
            CheckIoInfo = true;
            CheckCsvPattenList = true;
            CheckDirTimeSet = true;
            HasTimeSet = true;
            HasDirPatSet = true;

            Ac = true;
            Level = true;
        }
    }
}
