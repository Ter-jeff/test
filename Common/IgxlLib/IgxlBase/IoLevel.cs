namespace IgxlLib.IgxlBase
{
    public class IoLevel(string pinName, string vil, string vih, string vol, string voh, string vohAlt1, string vohAlt2, string iol, string ioh, string vt, string vcl, string vch, string vOutLoTyp, string vOutHiTyp, string driverMode)
    {
        #region Property
        public string PinName { get; set; } = pinName;
        public string Vil { get; set; } = vil;
        public string Vih { get; set; } = vih;
        public string Vol { get; set; } = vol;
        public string Voh { get; set; } = voh;
        public string VohAlt1 { get; set; } = vohAlt1;
        public string VohAlt2 { get; set; } = vohAlt2;
        public string Iol { get; set; } = iol;
        public string Ioh { get; set; } = ioh;
        public string Vt { get; set; } = vt;
        public string Vcl { get; set; } = vcl;
        public string Vch { get; set; } = vch;
        public string VOutLoTyp { get; set; } = vOutLoTyp;
        public string VOutHiTyp { get; set; } = vOutHiTyp;
        public string DriverMode { get; set; } = driverMode;

        #endregion
    }
}
