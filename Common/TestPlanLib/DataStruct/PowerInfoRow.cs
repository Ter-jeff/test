using System.Text.RegularExpressions;

using CommonReaderLib;

namespace TestPlanLib.DataStruct
{
    public partial class PowerInfoRow : MyRow
    {
        [GeneratedRegex(@"t\d+:init.*", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"t\d+:initlo.*", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"t\d+:inithi.*", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        [GeneratedRegex(@"t\d+:init.*", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex3();
        [GeneratedRegex(@"^\d+", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex4();
        private static readonly Regex _regex = MyRegex();
        private static readonly Regex _regex1 = MyRegex();
        private static readonly Regex _regex2 = MyRegex1();
        private static readonly Regex _regex3 = MyRegex2();
        private static readonly Regex _regex4 = MyRegex1();
        private static readonly Regex _regex5 = MyRegex2();
        private static readonly Regex _regex6 = MyRegex3();
        private static readonly Regex _regex7 = MyRegex3();
        private static readonly Regex _regex8 = MyRegex4();

        public string Chiplet { get; set; } = string.Empty;
        public string PinName { get; set; } = string.Empty;
        public string PowerSequence { get; set; } = string.Empty;
        public string PowerDownSequence { get; set; } = string.Empty;
        public string Ifold { get; set; } = string.Empty;
        public bool IsPowerPin
        {
            get { return _regex8.IsMatch(PowerSequence); }
        }
        public bool IsIoPowerPin
        {
            get { return !_regex7.IsMatch(PowerSequence); }
        }
        public bool IsIoPowerDownPin
        {
            get { return !_regex6.IsMatch(PowerDownSequence); }
        }
        public bool IsPowerPinHi
        {
            get { return _regex5.IsMatch(PowerSequence); }
        }
        public bool IsPowerPinLo
        {
            get { return _regex4.IsMatch(PowerSequence); }
        }
        public bool IsPowerDownPinHi
        {
            get { return _regex3.IsMatch(PowerDownSequence); }
        }
        public bool IsPowerDownPinLo
        {
            get { return _regex2.IsMatch(PowerDownSequence); }
        }
        public bool IsInitPin
        {
            get
            {
                return _regex1.IsMatch(PowerSequence) ||
                    _regex.IsMatch(PowerDownSequence);
            }
        }

        public PowerInfoRow(PowerInfoRow powerInfoRow)
        {
            if (powerInfoRow == null)
            {
                return;
            }

            Chiplet = powerInfoRow.Chiplet;
            PinName = powerInfoRow.PinName;
            PowerSequence = powerInfoRow.PowerSequence;
            PowerDownSequence = powerInfoRow.PowerDownSequence;
            Ifold = powerInfoRow.Ifold;
        }

        public PowerInfoRow()
        {
        }

        public PowerInfoRow Copy()
        {
            return new PowerInfoRow(this);
        }
    }
}
