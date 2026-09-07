using System.Collections.Generic;

using BinCutScriptLib.Base.Line;

namespace BinCutScriptLib.Algorithm.GradeSearch
{
    //[INFO]  ==================================================
    //[INFO] Creating overlay Overlay_BV_AMPLP5X_EXTSHSRXMD1CAPI_PP_HIDB0_S_FULP_AN_DDIO_ELB_JTG_LP5_MDX001_SI_EXTSHSRX_MD1CAPI_CS_LV
    //[INFO][Site 0] Spec voltage: VDD_CIO_VOP_VAR = 0.76 V
    //[INFO][Site 1] Spec voltage: VDD_CIO_VOP_VAR = 0.76 V
    //[INFO][Site 2] Spec voltage: VDD_CIO_VOP_VAR = 0.76 V
    //[INFO][Site 3] Spec voltage: VDD_CIO_VOP_VAR = 0.76 V
    //[INFO][Site 0] Spec voltage: VDD_CPU_SRAM_VOP_VAR = 0.68 V
    //[INFO][Site 1] Spec voltage: VDD_CPU_SRAM_VOP_VAR = 0.68 V
    //[INFO][Site 2] Spec voltage: VDD_CPU_SRAM_VOP_VAR = 0.68 V
    //[INFO][Site 3] Spec voltage: VDD_CPU_SRAM_VOP_VAR = 0.68 V
    //[INFO] Created overlay Overlay_BV_AMPLP5X_EXTSHSRXMD1CAPI_PP_HIDB0_S_FULP_AN_DDIO_ELB_JTG_LP5_MDX001_SI_EXTSHSRX_MD1CAPI_CS_LV
    //[INFO]  ==================================================
    internal class PayLoad
    {
        public int Site { get; set; }
        public string Pin { get; set; } = string.Empty;
        public double Voltage { get; set; }
        public List<BinCutLineBase> PayLoadLine { get; set; } = [];
        public string Line { get; set; } = string.Empty;
        public int LineNo { get; set; }
    }
}
