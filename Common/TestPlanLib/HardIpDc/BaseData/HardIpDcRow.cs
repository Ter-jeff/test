using System;

using CommonLib.Extension;

namespace TestPlanLib.HardIpDc.BaseData
{
    public class HardIpDcRow
    {
        private string _pinName = "";

        public string PinName
        {
            set
            {
                _pinName = value;
                if (_pinName.StartsWithIgnoreCase("VDD"))
                {
                    PinType = EnumHardIpDcPinType.Power;
                }
                else if (_pinName.StartsWithIgnoreCase("Pins"))
                {
                    PinType = EnumHardIpDcPinType.LevelIo;
                }
                else if (_pinName.Contains("DIFF", StringComparison.OrdinalIgnoreCase))
                {
                    PinType = EnumHardIpDcPinType.IoDiff;
                }
                else
                {
                    PinType = EnumHardIpDcPinType.IoSingle;
                }
            }
            get { return _pinName; }
        }
        public string Nv { set; get; }
        public string NvValt { set; get; }
        public string LvRatio { set; get; }
        public string HvRatio { set; get; }
        public string Ids { set; get; }
        public string Ifold { set; get; }
        public string Vil { set; get; }
        public string Vih { set; get; }
        public string Vol { set; get; }
        public string Voh { set; get; }
        public string Iol { set; get; }
        public string Ioh { set; get; }
        public string Vt { set; get; }
        public string Vcl { set; get; }
        public string Vch { set; get; }
        public string DriverMode { set; get; }
        public string Vicm { set; get; }
        public string Vid { set; get; }
        public string Vod { set; get; }
        public string RowNum { set; get; } = "";

        public EnumHardIpDcPinType PinType { set; get; }
        public bool NeedRatio { set; get; }

        public HardIpDcRow()
        {
            PinName = "";
            Nv = "";
            NvValt = "";
            LvRatio = "";
            HvRatio = "";
            Ids = "";
            Ifold = "";
            Vil = "";
            Vih = "";
            Vol = "";
            Voh = "";
            Iol = "";
            Ioh = "";
            Vt = "";
            Vcl = "";
            Vch = "";
            DriverMode = "";
            Vicm = "";
            Vid = "";
            Vod = "";
            NeedRatio = false;
        }

        public string FindSpecValue(string specName)
        {
            return specName.ToUpper() switch
            {
                "VIL" => Vil,
                "VIH" => Vih,
                "VOL" => Vol,
                "VOH" => Voh,
                "VT" => Vt,
                "VICM" => Vicm,
                "VID" => Vid,
                "VOD" => Vod,
                _ => "",
            };
        }
    }
}
