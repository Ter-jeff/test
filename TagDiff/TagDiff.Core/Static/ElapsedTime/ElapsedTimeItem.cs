using System;

namespace TagDiff.Core.Static.ElapsedTime
{
    public class ElapsedTimeItem(string module, TimeSpan timeSpan)
    {
        public string SheetName { get; set; } = module;
        public TimeSpan Time { get; set; } = timeSpan;
    }
}
