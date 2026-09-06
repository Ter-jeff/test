using System.Collections.Generic;

using BinCutScriptLib.Base.Line;

namespace BinCutScriptLib.Reader
{
    internal class ReadProductCsRow
    {
        public int Number { get; set; }
        public int Site { get; set; }
        public string TestName { get; set; } = string.Empty;
        public string TestBinCut { get; set; } = string.Empty;
        public string Pin { get; set; } = string.Empty;
        public string Channel { get; set; } = string.Empty;
        public string Low { get; set; } = string.Empty;
        public string LowUnit { get; set; } = string.Empty;
        public string Measured { get; set; } = string.Empty;
        public string MeasuredUnit { get; set; } = string.Empty;
        public string JudgeFail { get; set; } = string.Empty;
        public string High { get; set; } = string.Empty;
        public string HighUnit { get; set; } = string.Empty;
        public string Force { get; set; } = string.Empty;
        public int Loc { get; set; }
        public BinCutLineBase Line { get; internal set; } = new BinCutLineBase();

        internal string Print()
        {
            var texts = new List<string>()
            {
                Number.ToString(),
                Site.ToString(),
                TestName,
                TestBinCut,
                Pin,
                Channel,
                Low,
                LowUnit,
                Measured,
                MeasuredUnit,
                JudgeFail,
                High,
                HighUnit,
                Force,
                Loc.ToString(),
            };
            return string.Join("\t", texts);
        }
    }
}
