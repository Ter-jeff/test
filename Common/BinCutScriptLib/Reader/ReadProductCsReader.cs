using System;
using System.Collections.Generic;

using BinCutScriptLib.Base.Line;

using CommonLib.Extension;

namespace BinCutScriptLib.Reader
{
    internal class ReadProductCsReader(List<BinCutLineBase> binCutLineBases)
    {
        public List<ReadProductCsRow> ReadProductCsRows = Read(binCutLineBases);

        private static List<ReadProductCsRow> Read(List<BinCutLineBase> binCutLineBases)
        {
            var readProductCsRows = new List<ReadProductCsRow>();
            foreach (BinCutLineBase line in binCutLineBases)
            {
                if (line.LineNo == 1200)
                {

                }
                if (line.Line.StartsWithIgnoreCase("[WARN] ") ||
                    line.Line.Length == 0 ||
                    line.Line.Contains('<') ||
                    line.Line.Contains("Test Name") ||
                    !line.Line.Contains("Bincut"))
                {
                    continue;
                }

                string[] parts = line.Line.Split([' '], StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length > 12)
                {
                    if (parts.Length >= 13 && (line.Line.Contains("(F)") || line.Line.Contains("(A)")))
                    {
                        var record = new ReadProductCsRow
                        {
                            Number = int.Parse(parts[0]),
                            Site = int.Parse(parts[1]),
                            TestName = parts[2],
                            TestBinCut = parts[3],
                            Channel = parts[4],
                            Pin = "",
                            Low = parts[5],
                            LowUnit = parts[6],
                            Measured = parts[7],
                            MeasuredUnit = parts[8],
                            JudgeFail = parts[9],
                            High = parts[10],
                            HighUnit = parts[11],
                            Force = parts[12],
                            Loc = int.Parse(parts[13]),
                            Line = line
                        };
                        readProductCsRows.Add(record);
                    }
                    else
                    {
                        var record = new ReadProductCsRow
                        {
                            Number = int.Parse(parts[0]),
                            Site = int.Parse(parts[1]),
                            TestName = parts[2],
                            TestBinCut = parts[3],
                            Channel = parts[4],
                            Pin = "",
                            Low = parts[5],
                            LowUnit = parts[6],
                            Measured = parts[7],
                            MeasuredUnit = parts[8],
                            JudgeFail = " ",
                            High = parts[9],
                            HighUnit = parts[10],
                            Force = parts[11],
                            Loc = int.Parse(parts[12]),
                            Line = line
                        };
                        readProductCsRows.Add(record);
                    }
                }
                else
                {
                    continue;
                }
            }
            return readProductCsRows;
        }
    }
}
