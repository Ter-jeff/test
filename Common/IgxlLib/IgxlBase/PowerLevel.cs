namespace IgxlLib.IgxlBase
{
    public class PowerLevel
    {
        public string PinName { get; set; }
        public string Vmain { get; set; }
        public string Valt { get; set; }
        public string IFoldLevel { get; set; }
        public string Tdelay { get; set; }
        public string VSlewRate { get; set; } = string.Empty;
        public string Comment { get; set; }

        public PowerLevel(string pinName, string vmain, string valt, string iFoldLevel, string tdelay, string comment = "")
        {
            PinName = pinName;
            Vmain = vmain;
            Valt = valt;
            IFoldLevel = iFoldLevel;
            Tdelay = tdelay;
            Comment = comment;
        }

        public PowerLevel(string pinName, string vmain, string valt, string iFoldLevel, string tdelay, string vSlewRate, string comment = "")
        {
            PinName = pinName;
            Vmain = vmain;
            Valt = valt;
            IFoldLevel = iFoldLevel;
            Tdelay = tdelay;
            VSlewRate = vSlewRate;
            Comment = comment;
        }
    }
}
