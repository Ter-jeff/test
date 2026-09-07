using System.Linq;

using CommonLib.Extension;

namespace IgxlLib.Utility.Tester
{
    public class TesterConfig
    {
        public string Vsm = string.Empty;
        public string Io = string.Empty;
        public string HexVs = string.Empty;
        public string Uvs64 = string.Empty;
        public string Uvs256 = string.Empty;
        public string Uvs256Hp = string.Empty;
        public string Uvi80 = string.Empty;
        public string UltraPac = string.Empty;
        public string Dc30 = string.Empty;
        public string Us10G = string.Empty;
        public string Uw24Source = string.Empty;
        public string Uw24Measure = string.Empty;
        public string Mx8 = string.Empty;
        public string Support = string.Empty;

        public string GetToolType(string ch)
        {

            if (Vsm.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "VSM";
            }

            if (Io.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "I/O";
            }

            if (HexVs.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "HexVS";
            }

            if (Uvs256.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "UVS256";
            }

            if (Uvs256Hp.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "UVS256HP";
            }

            if (Uvs64.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "UVS64";
            }

            if (Uvi80.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "UVI80";
            }

            if (UltraPac.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "UltraPAC";
            }

            if (Dc30.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "DC30";
            }

            if (Us10G.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "US10G";
            }

            if (Support.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "Support";
            }

            if (Uw24Source.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "UW24Source";
            }

            if (Uw24Measure.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "UW24Measure";
            }

            if (Mx8.Split(',').Any(x => x.EqualsIgnoreCase(ch)))
            {
                return "MX8";
            }

            return "";
        }
    }
}
