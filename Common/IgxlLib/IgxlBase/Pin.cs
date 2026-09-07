using System.Diagnostics;

namespace IgxlLib.IgxlBase
{
    [DebuggerDisplay("{PinName}")]
    public class Pin : PinBase
    {
        public string Comment { get; set; } = string.Empty;

        public Pin()
        {
        }

        public Pin(string pinName, string pinType, string comment = "") : base(pinName, pinType)
        {
            Comment = comment;
        }

        public Pin(Pin pin) : base(pin)
        {
            if (pin == null)
            {
                return;
            }

            Comment = pin.Comment;
        }

        public Pin Copy()
        {
            return new Pin(this);
        }
    }
}
