using System;
using System.Collections.Generic;

using BinCutScriptLib.Base.Line;

using CommonLib.Extension;

namespace BinCutScriptLib.Reader
{
    internal class HarvestReader(List<BinCutLineBase> binCutLineBases)
    {
        public List<HarvestRow> Rows = Read(binCutLineBases);

        private static List<HarvestRow> Read(List<BinCutLineBase> binCutLineBases)
        {
            var rows = new List<HarvestRow>();
            foreach (BinCutLineBase line in binCutLineBases)
            {
                if (line.Line.StartsWithIgnoreCase("[WARN] "))
                {
                    continue;
                }

                string[] arr = line.Line.Split([" K ", " "], StringSplitOptions.RemoveEmptyEntries);
                if (arr.Length < 4)
                {
                    continue;
                }

                _ = arr[2];
                int channelIndex = CommonLib.Datalog.LineBase.GetChannelIndex(arr);
                int measureIndex = CommonLib.Datalog.LineBase.GetMeasureIndex(channelIndex, arr);
                if (!int.TryParse(arr[0], out int number) && rows.Count > 0)
                {
                    break;
                }

                _ = int.TryParse(arr[1], out int _);
                var row = new HarvestRow
                {
                    Number = number,
                    Site = int.Parse(arr[1]),
                    TestName = arr[2],
                    Pin = arr[3],
                    Channel = arr[channelIndex],
                    Measured = arr[measureIndex],
                };
                rows.Add(row);
            }
            return rows;
        }
    }
}
