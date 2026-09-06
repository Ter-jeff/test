using System;

namespace BinCutScriptLib.Base.Line
{
    public class BvEqLine : BinCutLineBase
    {
        public void GetbvNameAndSite(ref string bvName, ref int site)
        {
            //BV_VDD_SOC_MS001,2,VDD_CPU=0.543,VDD_GPU=0.596,VDD_SOC=0.606,VDD_CPU_SRAM=0.750,VDD_GPU_SRAM=0.668,VDD_FIXED=0.843,VDD_LOW=0.746
            string[] spt = Line.Split([','], StringSplitOptions.RemoveEmptyEntries);
            if (spt.Length > 2)
            {
                bvName = spt[0];
                if (!int.TryParse(spt[1], out site))
                {
                    throw new Exception($"Site errot @ {LineNo}: {Line}");
                }
            }
        }
    }
}
