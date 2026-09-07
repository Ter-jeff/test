using System;
using System.Collections.Generic;

using BinCutScriptLib.Base.Line;

using CommonLib.Extension;

namespace BinCutScriptLib.Reader
{
    internal class JudgestoredIdsReader(List<BinCutLineBase> binCutLineBases)
    {
        public List<JudgestoredIdsRow> JudgestoredIdsRows = Read(binCutLineBases);

        private static List<JudgestoredIdsRow> Read(List<BinCutLineBase> binCutLineBases)
        {
            var judgestoredIdsRows = new List<JudgestoredIdsRow>();
            foreach (BinCutLineBase line in binCutLineBases)
            {
                string[] arr = line.Line.Trim().Split(' ');

                if (line.Line.StartsWithIgnoreCase("[WARN] ") || !long.TryParse(arr[0], out _))
                {
                    continue;
                }

                string[] parts = line.Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 14 && (line.Line.Contains("(F)") || line.Line.Contains("(A)")))
                {
                    var record = new JudgestoredIdsRow
                    {
                        Number = int.Parse(parts[0]),
                        Site = int.Parse(parts[1]),
                        TestName = parts[2],
                        TestBinCutBin = parts[3],
                        TestNameIds = parts[4],
                        Pin = parts[5],
                        Channel = parts[6],
                        Low = parts[7],
                        LowUnit = parts[8],
                        Measured = parts[9],
                        MeasuredUnit = parts[10],
                        JudgeFail = parts[11],
                        High = parts[12],
                        HighUnit = parts[13],
                        Force = parts[14],
                        Loc = int.Parse(parts[15]),
                        Line = line
                    };
                    judgestoredIdsRows.Add(record);
                }
                else if (parts[5] == "-1")
                {
                    var record = new JudgestoredIdsRow
                    {
                        Number = int.Parse(parts[0]),
                        Site = int.Parse(parts[1]),
                        TestName = parts[2],
                        TestBinCutBin = parts[3],
                        TestNameIds = parts[4],
                        Pin = "",
                        Channel = parts[5],
                        Low = parts[6],
                        LowUnit = parts[7],
                        Measured = parts[8],
                        MeasuredUnit = parts[9],
                        JudgeFail = " ",
                        High = parts[10],
                        HighUnit = parts[11],
                        Force = parts[12],
                        Loc = int.Parse(parts[13]),
                        Line = line
                    };
                    judgestoredIdsRows.Add(record);
                }
                else
                {
                    var record = new JudgestoredIdsRow
                    {
                        Number = int.Parse(parts[0]),
                        Site = int.Parse(parts[1]),
                        TestName = parts[2],
                        TestBinCutBin = parts[3],
                        TestNameIds = parts[4],
                        Pin = parts[5],
                        Channel = parts[6],
                        Low = parts[7],
                        LowUnit = parts[8],
                        Measured = parts[9],
                        MeasuredUnit = parts[10],
                        JudgeFail = " ",
                        High = parts[11],
                        HighUnit = parts[12],
                        Force = parts[13],
                        Loc = int.Parse(parts[14]),
                        Line = line
                    };
                    judgestoredIdsRows.Add(record);
                }

            }
            return judgestoredIdsRows;
        }
    }
}
