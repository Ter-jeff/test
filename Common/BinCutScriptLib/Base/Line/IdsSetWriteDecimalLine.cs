using System;
using System.Text.RegularExpressions;

using BinCutScriptLib.Static;

namespace BinCutScriptLib.Base.Line
{
    public class IdsSetWriteDecimalLine : BinCutLineBase
    {
        public void GetIdsSetWriteDecimal(out string idsName, out double realVal, out int site, out double idsLsb)
        {
            int idsNameIdx, idsValIdx, lsbIdx;

            //Case1
            //Site(0)  CFGFuse IDS_SetWriteDecimal                    IDS_VDD_CPU = 251         (50.3197500567858 mA / 0.0002  )
            //Site(0)  CFGFuse IDS_SetWriteDecimal                    IDS_VDD_FIXED = 45         (9.00332593917847 mA / 0.0002  )

            //Case2  //new IDS format
            //Site(1)  CFGauto_eFuse_SetWriteVariable_SiteAware                             ids_vdd_gpu1_25c_7 = 201        (40.156751 mA / 0.200000mA)

            string[] spt = Line.Split([' ', '(', ')', '[', ']'], StringSplitOptions.RemoveEmptyEntries);
            if (Line.Contains("eFuse_SetWriteVariable_SiteAware"))
            { idsNameIdx = 3; idsValIdx = 6; lsbIdx = 9; }
            else if (Line.Contains("Set eFuse"))
            { idsNameIdx = 6; idsValIdx = 9; lsbIdx = 12; }
            else
            { idsNameIdx = 4; idsValIdx = 7; lsbIdx = 10; }

            idsName = spt[idsNameIdx].ToUpper();
            realVal = double.Parse(spt[idsValIdx]);
            site = int.Parse(spt[1]);
            Match allMatchs = Reg.RegexValue.Match(spt[lsbIdx]);
            idsLsb = 0.1;
            if (allMatchs.Length != 0)
            {
                idsLsb = double.Parse(allMatchs.Value);
            }
        }

        public void GetIdsSetWriteDecimalCs(out string idsName, out double realVal, out int site, out double idsLsb)
        {
            //Case1
            //[INFO]  [Site 0] Stored ids value in cache ids_vdd_cio = 1.4 (1.34253824448 / 0.1 * 0.1)
            //[INFO]  [Site 1] Stored ids value in cache ids_vdd_cio = 1.3 (1.276405699376 / 0.1 * 0.1)
            int idsNameIdx, idsValIdx, lsbIdx;
            string[] arr = Line.Replace("'", "").Split([' ', '(', ')', '[', ']'], StringSplitOptions.RemoveEmptyEntries);
            if (Line.Contains("Stored ids value in cache"))
            {
                idsNameIdx = 8;
                idsValIdx = 11;
                lsbIdx = 13;
            }
            else if (Line.Contains("Set eFuse"))
            {
                idsNameIdx = 6;
                idsValIdx = 9;
                lsbIdx = 12;
            }
            else
            {
                idsNameIdx = 8;
                idsValIdx = 11;
                lsbIdx = 13;
            }

            idsName = arr[idsNameIdx].ToUpper().Trim('\'');
            _ = double.TryParse(arr[idsValIdx], out realVal);
            _ = int.TryParse(arr[2], out site);
            if (!double.TryParse(arr[lsbIdx], out idsLsb))
            {
                idsLsb = 0.1;
            }
        }
    }
}
