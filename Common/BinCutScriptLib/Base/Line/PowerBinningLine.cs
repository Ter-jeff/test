using System;

using CommonLib.Extension;

namespace BinCutScriptLib.Base.Line
{
    public class PowerBinningLine : BinCutLineBase
    {
        public PowerBinningLineRow GetPowerBinningRow()
        {
            string[] spt = Line.Split([" K ", " "], StringSplitOptions.RemoveEmptyEntries);
            int regxIdx = -1;
            for (int i = 0; i < spt.Length; i++)
            {
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
                if (double.TryParse(spt[i], out _) || spt[i].EqualsIgnoreCase("N/A"))
                {
                    moreTwoStep++;
                    if (moreTwoStep == 2)
                    {
                        step = i;
                        break;
                    }
                }
            }

            var powerBinningLineRow = new PowerBinningLineRow();
            _ = int.TryParse(spt[1], out powerBinningLineRow.Site);
            powerBinningLineRow.PowerName = spt[2];
            powerBinningLineRow.Type = spt[3];

            if (!double.TryParse(spt[step], out double value1))
            {
            }
            powerBinningLineRow.Value = value1;
            return powerBinningLineRow;
        }
    }

    public class PowerBinningLineRow
    {
        public int Site;
        public string PowerName = "";
        public string Type = "";
        public double LowLimit;
        public double Value;
        public double HighLimit;
    }
}
