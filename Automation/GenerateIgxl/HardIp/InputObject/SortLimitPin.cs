using System.Collections.Generic;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class SortLimitPin
    {
        public string PinName { get; set; }

        public string PinNameWoStage()
        {
            if (PinName.Contains("="))
            {
                return PinName.Split('=')[1];
            }

            return PinName;
        }

        public MeasPin MeasPinData { get; set; }

        private List<SortLimitPin> _useLimitPins;

        public List<SortLimitPin> UseLimitPins
        {
            set
            {
                _useLimitPins = value;
            }
            get
            {
                return _useLimitPins ?? (_useLimitPins = new List<SortLimitPin>());
            }
        }

        public void AddData(string pinName, MeasPin measPin)
        {
            UseLimitPins.Add(new SortLimitPin
            {
                PinName = pinName,
                MeasPinData = measPin
            });
        }
    }
}
