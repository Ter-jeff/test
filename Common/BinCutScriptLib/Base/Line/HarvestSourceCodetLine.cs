using System;

using CommonLib.Extension;

namespace BinCutScriptLib.Base.Line
{
    public class HarvestSourceCodetLine : BinCutLineBase
    {
        public HarvestSourceCodetRow GetHarvestSourceCodetRow()
        {
            //Site:0,F_GFX_CORE0=False, Src Bits = 34, HarvestSourceCode [ First(L) ==> Last(R) ] :0010010011000011001001001100100100
            //Site:1,F_GFX_CORE0=True, Src Bits = 34, HarvestSourceCode [ First(L) ==> Last(R) ] :0010110010101001001101000100110100
            //Site:2,F_GFX_CORE0=False, Src Bits = 34, HarvestSourceCode [ First(L) ==> Last(R) ] :0010010011000011001001001100100100
            //Site:3,F_GFX_CORE0=False, Src Bits = 34, HarvestSourceCode [ First(L) ==> Last(R) ] :0010010011000011001001001100100100
            //Site:1,.\PATTERN\PATX\GFX_SCAN\PP_HIDB0_L_INLP_SC_CFXX_TDF_COM_AUT_ALLFRV_SI_XORDSSC_2_B0_2407252019.PATX, Label = X10GR1, Src Bits = 24, HarvestSourceCode [ First(L) ==> Last(R) ] :100000000011111111110000
            var harvestSourceCodeRow = new HarvestSourceCodetRow();
            string[] spt = Line.Split(',');
            for (int i = 0; i < spt.Length; i++)
            {
                string item = spt[i].Trim();
                if (item.StartsWithIgnoreCase("SITE"))
                {
                    string[] spt2 = spt[i].Trim().ToUpper().Split(':');
                    _ = int.TryParse(spt2[1].Trim(), out int value);
                    harvestSourceCodeRow.Site = value;
                }
                else if (item.StartsWithIgnoreCase("Src Bits"))
                {
                    harvestSourceCodeRow.SrcBit = item;
                }
                else if (item.StartsWithIgnoreCase("HarvestSourceCode"))
                {
                    harvestSourceCodeRow.HarvestSourceCode = item;
                }
            }
            if (Line.Contains("TRUE", StringComparison.OrdinalIgnoreCase) || Line.Contains("FALSE", StringComparison.OrdinalIgnoreCase))
            {
                harvestSourceCodeRow.MainFlag = spt[1].Trim();
            }
            else
            {
                harvestSourceCodeRow.PatternName = spt[1].Trim();
            }

            harvestSourceCodeRow.Line = this;
            return harvestSourceCodeRow;
        }
    }

    public class HarvestSourceCodetRow
    {
        public int Site;
        public string MainFlag = string.Empty;
        public string SrcBit = string.Empty;
        public string HarvestSourceCode = string.Empty;
        public string PatternName = string.Empty;
        public string Assignment = string.Empty;
        public string Label = string.Empty;
        public HarvestSourceCodetLine Line = new();
    }
}
