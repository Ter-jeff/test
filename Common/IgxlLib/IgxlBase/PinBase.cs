using System.Diagnostics;

namespace IgxlLib.IgxlBase
{
    [DebuggerDisplay("{PinName}")]
    public abstract class PinBase : IgxlRow
    {
        public string PinName { get; set; } = string.Empty;
        public string PinType { get; set; } = string.Empty;
        public string ChannelType { get; set; } = string.Empty;
        public string InstrumentType { get; set; } = string.Empty;

        protected PinBase(string pinName, string pinType)
        {
            PinName = pinName;
            PinType = pinType;
        }

        protected PinBase()
        {
        }

        protected PinBase(PinBase pinBase) : base(pinBase)
        {
            if (pinBase == null)
            {
                return;
            }

            PinName = pinBase.PinName;
            PinType = pinBase.PinType;
            ChannelType = pinBase.ChannelType;
            InstrumentType = pinBase.InstrumentType;
        }
    }
}
