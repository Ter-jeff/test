using System.Collections.Generic;

using BinCutScriptLib.Base.Line;

namespace BinCutScriptLib.Algorithm.GradeSearch
{
    //[INFO] ======== Setup Dig Src Test Start ========
    //[INFO]  Pattern = PP_HIDA0_S_IN03_SC_CXXX_XXX_XXX_AUT_MSXXXX_SI_SRMDSSC_2_A0_2402191849
    //[INFO]  Src Bits = 4
    //[INFO]  SrcPin = JTAG_TDI
    //[INFO]  DataSequence:
    //[INFO]  Assignment:C=Selsram()
    //[INFO]  Output String [ LSB(L) ==> MSB(R) ]:
    //[INFO]  [Site 0] 0100(C)
    //[INFO]  [Site 1] 0100(C)
    //[INFO]  [Site 2] 0100(C)
    //[INFO]  [Site 3] 0100(C)
    //[INFO]  Pattern = PP_HIDB0_L_INLP_SC_CFXX_SAA_COM_AUT_ALLFRV_SI_XORDSSC_2_B0_2407252019
    //[INFO]  Src Bits = 24
    //[INFO]  SrcPin = JTAG_TDI
    //[INFO]  DataSequence:
    //[INFO]  Assignment: C = Dssc(X10GR1)
    //[INFO]  Output String[LSB(L) ==> MSB(R)]:
    //[INFO][Site 0] 100000000011111111110000(C)
    //[INFO][Site 1] 100000000011111111110000(C)
    //[INFO][Site 2] 100000000011111111110000(C)
    //[INFO][Site 3] 100000000011111111110000(C)
    //[INFO] ======== Setup Dig Src Test End   ========
    internal class DigSrcCs
    {
        public string Pattern { get; set; } = string.Empty;
        public int SrcBits { get; set; }
        public string SrcPin { get; set; } = string.Empty;
        public string DataSequence { get; set; } = string.Empty;
        public string Assignment { get; set; } = string.Empty;
        public string OutputString { get; set; } = string.Empty;
        public List<BinCutLineBase> OutputStrings { get; set; } = [];
    }
}
