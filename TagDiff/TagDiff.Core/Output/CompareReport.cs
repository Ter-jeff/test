using System.Diagnostics;

using IgxlLib.Enums;

namespace TagDiff.Core.Output
{
    [DebuggerDisplay("{SheetName}")]
    public class CompareReport
    {
        public string Category { get; set; } = string.Empty;

        // ReSharper disable once UnusedMember.Global
        public string SubCategory
        {
            get
            {
                return Category == Const.Category.Hardip || Category == Const.Category.RegAssign ? "Hip" : "non-Hip";
            }
        }
        public string Block { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public EnumSheetType SheetType;
        public string Link { get; set; } = string.Empty;
        public bool Exclude { get; set; }
        // ReSharper disable once UnusedMember.Global
        public decimal AutoGenRate
        {
            get => GetPercentage(TotalCount - ManualCount - SemiAutoCount, TotalCount);
        }
        // ReSharper disable once UnusedMember.Global
        public decimal ManualRate
        {
            get => GetPercentage(ManualCount + SemiAutoCount, TotalCount);
        }

        public int TotalCount { get; set; }
        public int PassCount;
        public int ManualCount { get; set; }
        public int SemiAutoCount { get; set; }
        // ReSharper disable once UnusedMember.Global
        public int FailCount
        {
            get => ManualCount + SemiAutoCount;
        }
        public string Description { get; private set; } = string.Empty;

        private static decimal GetPercentage(decimal left, decimal right)
        {
            return Divide(left, right);
        }

        private static decimal Divide(decimal left, decimal right)
        {
            if (right == 0)
            {
                return 0;
            }
            else
            {
                return decimal.Divide(left, right);
            }
        }

        internal void AddDescription(string text)
        {
            if (!string.IsNullOrEmpty(Description))
            {
                Description += ";" + text;
            }
            else if (!string.IsNullOrEmpty(text))
            {
                Description = text;
            }
        }
    }
}
