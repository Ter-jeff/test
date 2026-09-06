using System.Collections.Generic;
using System.Linq;

namespace IgxlLib.IgxlBase
{
    public class TSet : IgxlRow
    {
        public string Name { get; set; } = string.Empty;
        public string CyclePeriod { get; set; } = string.Empty;
        public List<TimingRow> TimingRows { get; set; } = [];

        public TSet()
        {
        }

        public TSet(TSet tSet) : base(tSet)
        {
            if (tSet == null)
            {
                return;
            }

            Name = tSet.Name;
            CyclePeriod = tSet.CyclePeriod;

            if (tSet.TimingRows != null)
            {
                TimingRows = [.. tSet.TimingRows.Select(row => row.Copy())];
            }
        }

        public void AddTimingRow(TimingRow timingRow)
        {
            TimingRows.Add(timingRow);
        }

        public TSet Copy()
        {
            return new TSet(this);
        }
    }
}
