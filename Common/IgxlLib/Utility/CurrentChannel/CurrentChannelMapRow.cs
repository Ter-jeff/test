using IgxlLib.IgxlBase;

namespace IgxlLib.Utility.CurrentChannel
{
    public class CurrentChannelMapRow : IgxlRow
    {
        public string TesterChannel { get; set; } = string.Empty;
        public string DibChannel { get; set; } = string.Empty;
        public string SignalName { get; set; } = string.Empty;

        public CurrentChannelMapRow()
        {
        }

        public CurrentChannelMapRow(CurrentChannelMapRow currentChannelMapRow) : base(currentChannelMapRow)
        {
            if (currentChannelMapRow == null)
            {
                return;
            }

            TesterChannel = currentChannelMapRow.TesterChannel;
            DibChannel = currentChannelMapRow.DibChannel;
            SignalName = currentChannelMapRow.SignalName;
        }

        public CurrentChannelMapRow Copy()
        {
            return new CurrentChannelMapRow(this);
        }
    }
}
