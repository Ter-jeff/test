namespace IgxlLib.IgxlBase
{
    public class DcviPowerLevel(string pinName, string vps, string isc, string delay, string comment)
    {
        public string PinName { get; set; } = pinName;
        public string Vps { get; set; } = vps;
        public string Isc { get; set; } = isc;
        public string TDelay { get; set; } = delay;
        public string Comment { get; set; } = comment;
    }
}
