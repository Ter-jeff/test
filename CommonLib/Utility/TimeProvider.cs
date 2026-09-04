using System;

namespace CommonLib.Utility
{
    public abstract class TimeProvider
    {
        public static TimeProvider Current { get; set; } = DefaultTimeProvider.Instance;

        public abstract DateTime Now { get; }

        public static void ResetToDefault()
        {
            Current = DefaultTimeProvider.Instance;
        }
    }

    public class DefaultTimeProvider : TimeProvider
    {
        private DefaultTimeProvider() { }

        public static DefaultTimeProvider Instance { get; } = new DefaultTimeProvider();

        public override DateTime Now => DateTime.Now;
    }
}
