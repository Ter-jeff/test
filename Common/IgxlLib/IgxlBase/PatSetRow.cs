using System;

using CommonLib.Extension;

namespace IgxlLib.IgxlBase
{
    public class PatSetRow : IgxlRow
    {
        public string PatternSet { get; set; } = string.Empty;
        public string TdGroup { get; set; } = string.Empty;
        public string TimeDomain { get; set; } = string.Empty;
        public string Enable { get; set; } = string.Empty;
        public string File { get; set; } = string.Empty;
        public string Burst { get; set; } = string.Empty;
        public string StartLabel { get; set; } = string.Empty;
        public string StopLabel { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;

        public bool CompareRow(PatSetRow patSetRow)
        {
            StringComparer comparer = StringExtensions.IgnoreCase;

            return comparer.Equals(TdGroup, patSetRow.TdGroup) &&
                comparer.Equals(TimeDomain, patSetRow.TimeDomain) &&
                comparer.Equals(Enable, patSetRow.Enable) &&
                comparer.Equals(File, patSetRow.File) &&
                comparer.Equals(Burst, patSetRow.Burst) &&
                comparer.Equals(Comment, patSetRow.Comment);
        }
    }
}
