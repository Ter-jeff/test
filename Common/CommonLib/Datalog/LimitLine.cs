using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CommonLib.Datalog
{
    public partial class LimitLine : LineBase
    {
        [GeneratedRegex(@"^Site\(\d+\) EFUSE (Write|Read) Values", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexEfuseLine();

        public static readonly Regex RegEfuseLine = MyRegexEfuseLine();
        private static readonly Dictionary<string, double> _unit = new()
        {
            { "K", 1000},
            { "M", 1000000},
            { "G", 1000000000},
            { "T", 1000000000000},
        };
        //Number   Site  Test Name                      Pin                       Channel  Low            Measured           High           Force          Loc     
        //50138538 0     IDS                            VDD_SOC                   11.x404  0.0000 A       26.0000 mA         103.2000 mA    0.0000         0       
        //50138538 1     IDS                            VDD_SOC                   21.x404  0.0000 A       22.0000 mA         103.2000 mA    0.0000         0       
        //50138538 2     IDS                            VDD_SOC                   11.x315  0.0000 A       26.4000 mA         103.2000 mA    0.0000         0     
        public string GetTestName()
        {
            string[] spt = Line.Split([" ", "(F)", "(A)"], StringSplitOptions.RemoveEmptyEntries);
            if (spt.Length > 2)
            {
                return spt[2].ToUpper();
            }

            return "";
        }

        public void GetSiteData(out int site, out double log)
        {
            log = -1;
            site = -1;
            string[] spt = Line.Split([" ", "(F)", "(A)"], StringSplitOptions.RemoveEmptyEntries);
            int channelIndex = GetChannelIndex(spt);
            if (channelIndex == 0)
            {
                channelIndex = Array.FindIndex(spt, str => str.Trim() == "-1");
            }
            if (channelIndex == -1)
            {
                return;
            }

            int measureIndex = GetMeasureIndex(channelIndex, spt);
            if (measureIndex == -1)
            {
                return;
            }

            _ = int.TryParse(spt[1], out site);
            log = double.Parse(spt[measureIndex]);
            if (spt[measureIndex + 1] == "uA")
            {
                log /= 1000.0;
            }
            else if (spt[measureIndex + 1] == "V")
            {
                log *= 1000.0;
            }
            else if (spt[measureIndex + 1] == "uA")
            {
                log /= 1000.0;
            }
            else if (spt[measureIndex + 1] == "A")
            {
                log *= 1000.0;
            }
        }

        public void GetSiteOnly(out int site)
        {
            string[] spt = Line.Split([" ", "(F)", "(A)"], StringSplitOptions.RemoveEmptyEntries);
            if (spt.Length < 2 || !int.TryParse(spt[1], out site))
            {
                site = -1;
            }
        }

        public bool IsFail()
        {
            return Line.Contains("(F)");
        }

        public LimitRow? ToRow()
        {
            if (Line.StartsWith("[INFO]") || RegEfuseLine.IsMatch(Line) || string.IsNullOrEmpty(Line))
            {
                return null;
            }

            string[] spt = Line.Split([" K ", " M ", " G ", " T ", " "], StringSplitOptions.RemoveEmptyEntries);
            int regexIdx = -1;
            for (int i = 0; i < spt.Length; i++)
            {
                Match matchObj = RegexChannel.Match(spt[i]);
                if (matchObj.Length != 0)
                {
                    regexIdx = i;
                    break;
                }
                if (spt[i] == "-1")
                {
                    regexIdx = i;
                    break;
                }
            }
            int moreTwoStep = 0;
            int step = 0;
            for (int i = regexIdx + 1; i < spt.Length; i++)
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

            var limitRow = new LimitRow
            {
                Number = spt[0]
            };
            _ = int.TryParse(spt[1], out limitRow.Site);
            limitRow.TestName = spt[2];
            limitRow.Pin = spt[3];

            if (double.TryParse(spt[step], out double value1))
            {
                limitRow.Measured = value1;
            }
            if (double.TryParse(spt[step - 1], out double value2))
            {
                limitRow.LowLimit = value2;
            }

            if (double.TryParse(spt[step + 1], out double value3))
            {
                limitRow.HighLimit = value3;
            }

            return limitRow;
        }

        public static void GetExecuteUnitVal(string line, string[] spt, int step, out double limitLow, out double limitHigh)
        {
            double unitValue = 1;
            if (line.Contains(" K "))
            {
                _unit.TryGetValue("K", out unitValue);
            }
            else if (line.Contains(" M "))
            {
                _unit.TryGetValue("M", out unitValue);
            }
            else if (line.Contains(" G "))
            {
                _unit.TryGetValue("G", out unitValue);
            }
            else if (line.Contains(" T "))
            {
                _unit.TryGetValue("T", out unitValue);
            }

            if (double.TryParse(spt[step - 1], out double valueLow))
            {
                limitLow = valueLow * unitValue;
            }
            else
            {
                limitLow = double.NaN;
            }
            if (double.TryParse(spt[step + 1], out double valueHigh))
            {
                limitHigh = valueHigh * unitValue;
            }
            else
            {
                limitHigh = double.NaN;
            }
        }
    }
}
