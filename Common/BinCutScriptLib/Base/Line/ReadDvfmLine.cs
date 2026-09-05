using System;
using System.Text.RegularExpressions;

using BinCutScriptLib.Static;

namespace BinCutScriptLib.Base.Line
{
    //<Read_DVFM_To_GradeVDD>
    //16351000 1     PASSBIN                                                                                                                                                 -1       1              1                  3              0              0       
    //16351000 2     PASSBIN                                                                                                                                                 -1       1              1                  3              0              0       
    //16351000 3     PASSBIN                                                                                                                                                 -1       1              1                  3              0              0       
    //16351001 1     VDD_PCPU_MC601 VDD Grade                                                                                                                                -1       N/A            646.8750 mV        N/A            0.0000 V       0       
    //16351002 1     VDD_PCPU_MC601 VDD Product                                                                                                                                                  -1       1              7                    8              0              0       
    public class ReadDvfmLine : BinCutLineBase
    {
        public ReadDvfmRow GetReadDvfmLineRow()
        {
            string[] spt = Line.Split([" K ", " "], StringSplitOptions.RemoveEmptyEntries);
            int regxIdx = -1;
            for (int i = 0; i < spt.Length; i++)
            {
                Match matchObj = Reg.RegxValue2.Match(spt[i]);
                if (matchObj.Length != 0)
                {
                    regxIdx = i;
                    break;
                }
                if (spt[i] == "-1")
                {
                    regxIdx = i;
                    break;
                }
            }
            int moreTwoStep = 0;
            int step = 0;
            for (int i = regxIdx + 1; i < spt.Length; i++)
            {
                if (spt[i] == "N/A" || double.TryParse(spt[i], out _))
                {
                    moreTwoStep++;
                    if (moreTwoStep == 2)
                    {
                        step = i;
                        break;
                    }
                }
            }

            var row = new ReadDvfmRow();
            _ = int.TryParse(spt[1], out row.Site);
            row.Name = spt[2];
            row.Type = spt[3];
            _ = double.TryParse(spt[step], out double value1);
            row.Measured = value1;
            row.Line = this;
            return row;
        }
    }

    public class ReadDvfmRow
    {
        public int Site;
        public string Name = "";
        public string Type = "";
        public double LowLimit;
        public double Measured;
        public double HighLimit;
        public BinCutLineBase Line = new();
    }
}
