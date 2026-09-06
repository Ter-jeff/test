using System.Collections.Generic;

using BinCutScriptLib.Base.Line;

namespace BinCutScriptLib.Reader
{
    internal class JudgestoredIdsRow
    {
        public int Number { get; set; }
        public int Site { get; set; }
        public string TestName { get; set; } = string.Empty;
        public string TestBinCutBin { get; set; } = string.Empty;
        public string TestNameIds { get; set; } = string.Empty;
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

        public JudgestoredIdsRow()
        {
        }

        public JudgestoredIdsRow(JudgestoredIdsRow judgestoredIdsRow)
        {
            if (judgestoredIdsRow == null)
            {
                return;
            }

            Number = judgestoredIdsRow.Number;
            Site = judgestoredIdsRow.Site;
            TestName = judgestoredIdsRow.TestName;
            TestBinCutBin = judgestoredIdsRow.TestBinCutBin;
            TestNameIds = judgestoredIdsRow.TestNameIds;
            Pin = judgestoredIdsRow.Pin;
            Channel = judgestoredIdsRow.Channel;
            Low = judgestoredIdsRow.Low;
            LowUnit = judgestoredIdsRow.LowUnit;
            Measured = judgestoredIdsRow.Measured;
            MeasuredUnit = judgestoredIdsRow.MeasuredUnit;
            JudgeFail = judgestoredIdsRow.JudgeFail;
            High = judgestoredIdsRow.High;
            HighUnit = judgestoredIdsRow.HighUnit;
            Force = judgestoredIdsRow.Force;
            Loc = judgestoredIdsRow.Loc;
            Line = judgestoredIdsRow.Line == null ? new BinCutLineBase() : new BinCutLineBase { Line = judgestoredIdsRow.Line.Line, LineNo = judgestoredIdsRow.Line.LineNo };
        }

        public JudgestoredIdsRow Copy()
        {
            return new JudgestoredIdsRow(this);
        }

        internal string Print()
        {
            var texts = new List<string>()
            {
                Number.ToString(),
                Site.ToString(),
                TestName,
                TestBinCutBin,
                TestNameIds,
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
