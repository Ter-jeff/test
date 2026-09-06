namespace IgxlLib.IgxlBase
{
    public class DiffLevel(string pinName, string vicm, string vid, string dVid0, string dVid1, string dVicm0, string dVicm1, string vod, string vodAlt1, string vodAlt2, string dVod0, string dVod1, string iol, string ioh, string vodTyp, string vocmTyp, string vt, string vcl, string vch, string driverMode)
    {
        public string PinName { get; set; } = pinName;
        public string Vicm { get; set; } = vicm;
        public string Vid { get; set; } = vid;
        public string DVid0 { get; set; } = dVid0;
        public string DVid1 { get; set; } = dVid1;
        public string DVicm0 { get; set; } = dVicm0;
        public string DVicm1 { get; set; } = dVicm1;
        public string Vod { get; set; } = vod;
        public string VodAlt1 { get; set; } = vodAlt1;
        public string VodAlt2 { get; set; } = vodAlt2;
        public string DVod0 { get; set; } = dVod0;
        public string DVod1 { get; set; } = dVod1;
        public string Iol { get; set; } = iol;
        public string Ioh { get; set; } = ioh;
        public string VodTyp { get; set; } = vodTyp;
        public string VocmTyp { get; set; } = vocmTyp;
        public string Vt { get; set; } = vt;
        public string Vcl { get; set; } = vcl;
        public string Vch { get; set; } = vch;
        public string DriverMode { get; set; } = driverMode;
    }
}
