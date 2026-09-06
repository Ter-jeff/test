using System;
using System.Diagnostics;

using TagDiff.Core.Static.ElapsedTime;

namespace TagDiff.Core.Utility
{
    public static class StopwatchMeasureSection
    {
        public static ElapsedTimeItem MeasureSection(string name, Action action)
        {
            ArgumentNullException.ThrowIfNull(action);

            var sw = Stopwatch.StartNew();
            try
            {
                action();
            }
            finally
            {
                sw.Stop();
            }

            return new ElapsedTimeItem(name ?? string.Empty, sw.Elapsed);
        }
    }
}
