using System.Collections.Generic;
using System.Diagnostics;

using CommonLib.Extension;

namespace IgxlLib.IgxlBase
{
    [DebuggerDisplay("{PinName}")]
    public class PinGroup : PinBase
    {
        public List<Pin> PinList { get; set; }

        public PinGroup(string pinGrpName, string pinType) : base(pinGrpName, pinType)
        {
            PinList = [];
        }

        public PinGroup(string pinGrpName)
        {
            PinList = [];
            PinName = pinGrpName;
        }

        public void AddPin(Pin pin)
        {
            if (!PinList.Exists(a => a.PinName!.EqualsIgnoreCase(pin.PinName!)))
            {
                PinList.Add(pin);
            }
        }

        public void AddPin(string pin, string pinType = "")
        {
            if (!PinList.Exists(a => a.PinName!.EqualsIgnoreCase(pin)))
            {
                var newPin = new Pin(pin, pinType);
                PinList.Add(newPin);
            }
        }

        public void AddPins(List<Pin> pins)
        {
            foreach (Pin pin in pins)
            {
                if (!PinList.Exists(a => a.PinName!.EqualsIgnoreCase(pin.PinName!)))
                {
                    var newPin = new Pin(pin.PinName!, pin.PinType!, pin.Comment ?? "")
                    {
                        ColumnA = pin.ColumnA
                    };
                    PinList.Add(newPin);
                }
            }
        }

        public void AddPins(List<string> pins, string pinType = "")
        {
            foreach (string pin in pins)
            {
                if (!PinList.Exists(a => a.PinName!.EqualsIgnoreCase(pin)))
                {
                    var newPin = new Pin(pin, pinType);
                    PinList.Add(newPin);
                }
            }
        }
    }
}
