using System;

namespace CommonLib.Static
{
    public static class TimeContext
    {
        public static TimeProvider Current { get; set; } = TimeProvider.System;

        public static DateTime Now => Current.GetLocalNow().DateTime;

        public static void ResetToDefault()
        {
            Current = TimeProvider.System;
        }
    }
}
