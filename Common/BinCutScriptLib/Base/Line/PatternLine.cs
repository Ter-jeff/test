using System;

using TestPlanLib.PatternListCsvFile;

namespace BinCutScriptLib.Base.Line
{
    public class PatternLine : BinCutLineBase
    {
        //50043176     0     MF001_SocTd_MAF011_pp_jcda0_h_pl01_sc_cc97_tdf_com_aut_MXXXXX_SI_BV                                            PP_JCDA0_H_IN00_SC_XXXX_XXX_XXX_AUT_ALLFRV_SI_12_A0_2008111321                                                                       N/A               N/A                  
        //50043177     0     MF001_SocTd_MAF011_pp_jcda0_h_pl01_sc_cc97_tdf_com_aut_MXXXXX_SI_BV                                            PP_JCDA0_H_IN00_SC_CXXX_XXX_XXX_AUT_ALLFRV_SI_DSRMDSSC_5_A0_2007261419                                                               N/A               N/A                  
        public PatternRow GetDatalogPatternRow()
        {
            string[] spt = Line.Split([" "], StringSplitOptions.RemoveEmptyEntries);
            var patternRow = new PatternRow();
            if (spt.Length <= 4)
            {
                return patternRow;
            }

            _ = int.TryParse(spt[1], out patternRow.Site);
            patternRow.Number = spt[0];
            patternRow.TestName = spt[2];
            patternRow.PatternName = spt[3];
            patternRow.IsFail = Line.Contains("(F)") || Line.Contains("(A)");
            patternRow.GenericPatternName = spt[3].Contains('_')
                ? new PatternNameInfo(spt[3]).GenericName.ToUpper()
                : "";
            patternRow.PatternLine = this;
            return patternRow;
        }

    }

    public class PatternRow
    {
        public bool IsFail;
        public string Number = "";
        public int Site;
        public string TestName = "";
        public string PatternName = "";
        public string GenericPatternName = "";
        public PatternLine PatternLine = new();

        public PatternRow()
        {
        }

        public PatternRow(PatternRow patternRow)
        {
            if (patternRow == null)
            {
                return;
            }

            IsFail = patternRow.IsFail;
            Number = patternRow.Number;
            Site = patternRow.Site;
            TestName = patternRow.TestName;
            PatternName = patternRow.PatternName;
            GenericPatternName = patternRow.GenericPatternName;
            PatternLine = patternRow.PatternLine == null ? new PatternLine() : new PatternLine { Line = patternRow.PatternLine.Line, LineNo = patternRow.PatternLine.LineNo };
        }

        public PatternRow Copy()
        {
            return new PatternRow(this);
        }
    }
}
