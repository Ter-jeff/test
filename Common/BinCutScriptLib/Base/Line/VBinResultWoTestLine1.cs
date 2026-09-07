using System;
using System.Linq;

using BinCutScriptLib.Static;

using CommonLib.Extension;

namespace BinCutScriptLib.Base.Line
{
    public class VBinResultWoTestLine1 : BinCutLineBase
    {
        public VBinResultWoTestRow GetVBinResultWoTestRow()
        {
            var row = new VBinResultWoTestRow();
            //Site:0,VDD_PCPU_MP105 doesn't need to be tested(no keyword in Non_Binning_Rail), VDD_PCPU_MP105 follows voltage from its Allow_Equal mode: VDD_PCPU_MP00A
            //Site:2,VDD_PCPU_MP105 doesn't need to be tested(no keyword in Non_Binning_Rail), VDD_PCPU_MP105 follows voltage from its Allow_Equal mode: VDD_PCPU_MP00A         
            string[] spt = Line.Split(',');
            for (int i = 0; i < spt.Length; i++)
            {
                if (spt[i].StartsWithIgnoreCase("SITE"))
                {
                    string[] spt2 = spt[i].Trim().ToUpper().Split(':');
                    _ = int.TryParse(spt2[1].Trim(), out int value);
                    row.Site = value;
                }
                else if (spt[i].Trim().Contains("doesn't need to be tested", StringComparison.OrdinalIgnoreCase))
                {
                    row.PowerMode = spt[i].Trim().Split(' ')[0];
                    row.Mode = row.PowerMode.Split('_').Last();
                }
                else if (spt[i].Trim().Contains("Allow_Equal mode:", StringComparison.OrdinalIgnoreCase))
                {
                    row.AllowEqualPowerMode = spt[i].Split(':')[1].Split(' ')[1].Trim();
                    row.AllowEqualMode = row.AllowEqualPowerMode.Split('_').Last();
                }
                else if (spt[i].Trim().Contains("used", StringComparison.OrdinalIgnoreCase) &&
                         spt[i].Trim().Contains("voltage", StringComparison.OrdinalIgnoreCase))
                {
                    string eqnString = spt[i].Trim().Split(' ')[2].Replace("EQ", "");

                    _ = int.TryParse(eqnString, out row.SetEqn);
                }
            }
            return row;
        }

        public VBinResultWoTestRow GetVBinResultWoTestRowCs()
        {
            var row = new VBinResultWoTestRow();
            //[INFO]  [Site 0] Set Allow Equal mode MP00E to follow MP00D. No test is required for MP00E.
            //[INFO]  [Site 1] Set Allow Equal mode MP00E to follow MP00D. No test is required for MP00E.
            //[INFO]  [Site 1] Set single equation mode MPS001 to use E1. No test is required for MPS001.
            //[INFO]  [Site 3] Set single equation mode MPS001 to use E1. No test is required for MPS001.
            //[INFO]  [Site 0] Set force mode MG001 to use E1
            if (Line.Contains("Site ", StringComparison.OrdinalIgnoreCase))
            {
                _ = int.TryParse(Reg.RegexSite1.Match(Line).Groups["site"].ToString(), out row.Site);
            }
            if (Line.Contains("Set Allow Equal mode ", StringComparison.OrdinalIgnoreCase))
            {
                int index = Line.IndexOf("Set Allow Equal mode ", StringComparison.OrdinalIgnoreCase) + "Set Allow Equal mode ".Length;
                string[] arr = Line[index..].Split(' ');
                row.Mode = arr[0];
                row.PowerMode = arr[0];
                row.AllowEqualMode = arr[3].Trim('.');
                row.AllowEqualPowerMode = arr[3].Trim('.');
            }
            if (Line.Trim().Contains("Set single equation mode ", StringComparison.OrdinalIgnoreCase))
            {
                int index = Line.IndexOf("Set single equation mode ", StringComparison.OrdinalIgnoreCase) + "Set single equation mode ".Length;
                string[] arr = Line[index..].Split(' ');
                row.Mode = arr[0];
                row.PowerMode = arr[0];
                string eqnString = arr[3].Replace("E", "");

                _ = int.TryParse(eqnString, out row.SetEqn);
            }
            if (Line.Trim().Contains("Set force mode ", StringComparison.OrdinalIgnoreCase))
            {
                int index = Line.IndexOf("Set force mode ", StringComparison.OrdinalIgnoreCase) + "Set force mode ".Length;
                string[] arr = Line[index..].Split(' ');
                row.Mode = arr[0];
                row.PowerMode = arr[0];
                string eqnString = arr[3].Replace("E", "");

                _ = int.TryParse(eqnString, out row.SetEqn);
            }
            return row;
        }
    }

    public class VBinResultWoTestRow : BinCutLineBase
    {
        public int Site;
        public string Mode = string.Empty;
        public string AllowEqualMode = string.Empty;
        public string PowerMode = string.Empty;
        public string AllowEqualPowerMode = string.Empty;
        public int SetEqn;
    }
}
