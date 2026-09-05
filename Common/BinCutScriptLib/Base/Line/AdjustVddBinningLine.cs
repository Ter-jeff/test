using System;
using System.Text.RegularExpressions;

using BinCutScriptLib.Static;

namespace BinCutScriptLib.Base.Line
{
    public class AdjustVddBinningLine : BinCutLineBase
    {
        public AdjustVddBinningRow GetAdjustVddBinningRow()
        {
            string[] spt = Line.Split([" K ", " "], StringSplitOptions.RemoveEmptyEntries);
            int regxIdx = -1;
            for (int i = 0; i < spt.Length; i++)
            {
                Match matchObj = Reg.RegexValue1.Match(spt[i]);
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
                if (double.TryParse(spt[i], out _))
                {
                    moreTwoStep++;
                    if (moreTwoStep == 2)
                    {
                        step = i;
                        break;
                    }
                }
            }

            var adjustVddBinningRow = new AdjustVddBinningRow();
            _ = int.TryParse(spt[1], out adjustVddBinningRow.Site);
            adjustVddBinningRow.PowerName = spt[2];
            adjustVddBinningRow.Type = spt[3];

            _ = double.TryParse(spt[step], out double value1);
            adjustVddBinningRow.Unit = spt[regxIdx + 2];
            adjustVddBinningRow.Value = value1;
            adjustVddBinningRow.Line = this;
            return adjustVddBinningRow;
        }

        public AdjustVddBinningRow GetAdjustVddBinningRowCsharp()
        {
            string[] arr = Line.Split([" K ", " "], StringSplitOptions.RemoveEmptyEntries);

            _ = int.Parse(arr[1]);

            _ = arr[3];

            _ = arr[2];
            int channelIndex = GetChannelIndex(arr);
            int measureIndex = GetMeasureIndex(channelIndex, arr);
            _ = double.TryParse(arr[measureIndex], out double value);

            var adjustVddBinningRow = new AdjustVddBinningRow();
            _ = int.TryParse(arr[1], out adjustVddBinningRow.Site);
            adjustVddBinningRow.PowerName = arr[2];
            adjustVddBinningRow.Type = arr[3];
            if (!double.TryParse(arr[measureIndex + 1], out double _))
            {
                adjustVddBinningRow.Unit = arr[measureIndex + 1];
            }

            adjustVddBinningRow.Value = value;
            adjustVddBinningRow.Line = this;
            return adjustVddBinningRow;
        }
    }

    public class AdjustVddBinningRow
    {
        public int Site;
        public string PowerName = "";
        public string Type = "";
        public double LowLimit;
        public double Value;
        public double HighLimit;
        public int XCoor = -1;
        public int YCoor = -1;
        public string Unit = "";
        public BinCutLineBase Line = new();
    }
}
