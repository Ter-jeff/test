namespace RfLib.InstrumentSetup.InstrumentTypeData
{
    public class InstrumentTypeSetting
    {
        public string InstrumentType { get; set; } = "";
        public string PinCheck { get; set; } = "";
        public string PinExtraName { get; set; } = "";
        public double HighLimit { get; set; }
        public double LowLimit { get; set; }
        public string Path { get; set; } = "";
        public string PinType { get; set; } = "";
    }
}
