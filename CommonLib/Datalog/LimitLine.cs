using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CommonLib.Datalog
{
    [Serializable]
    public class LimitLine : LineBase
    {
        public static readonly Regex RegEfuseLine = new Regex(@"^Site\(\d+\) EFUSE (Write|Read) Values", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static public Dictionary<string, double> Unit = new Dictionary<string, double>
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
            string[] spt = Line.Split(new[] { " ", "(F)", "(A)" }, StringSplitOptions.RemoveEmptyEntries);
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
            string[] spt = Line.Split(new[] { " ", "(F)", "(A)" }, StringSplitOptions.RemoveEmptyEntries);
            int channelIndex = GetChannelIndex(spt);
            if (channelIndex == 0)
            {
                channelIndex = Array.FindIndex(spt, str => str.Trim().Equals("-1"));
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

            int.TryParse(spt[1], out site);
            log = measureIndex != -1 ? double.Parse(spt[measureIndex]) : -1;
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
            string[] spt = Line.Split(new[] { " ", "(F)", "(A)" }, StringSplitOptions.RemoveEmptyEntries);
            if (spt.Length < 2 || !int.TryParse(spt[1], out site))
            {
                site = -1;
            }
        }

        public bool IsFail()
        {
            return Line.IndexOf("(F)", StringComparison.Ordinal) != -1;
        }

        public LimitRow ToRow()
        {
            if (Line.StartsWith("[INFO]") || RegEfuseLine.IsMatch(Line) || Line == "")
            {
                return null;
            }

            string[] spt = Line.Split(new[] { " K ", " M ", " G ", " T ", " " }, StringSplitOptions.RemoveEmptyEntries);
            int regxIdx = -1;
            for (int i = 0; i < spt.Length; i++)
            {
                Match matchObj = RegxChannel.Match(spt[i]);
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

            var limitRow = new LimitRow
            {
                Number = spt[0]
            };
            int.TryParse(spt[1], out limitRow.Site);
            limitRow.TestName = spt[2];
            limitRow.Pin = spt[3];
            double.TryParse(spt[step], out double value1);
            double.TryParse(spt[step - 1], out double value2);
            double.TryParse(spt[step + 1], out double value3);
            limitRow.Measured = value1;
            limitRow.LowLimit = value2;
            limitRow.HighLimit = value3;
            return limitRow;
        }

        public void GetExecuteUnitVal(string line, string[] spt, int step, out double limitlow, out double limithigh)
        {
            double unitValue = 1;
            if (line.Contains(" K "))
            {
                Unit.TryGetValue("K", out unitValue);
            }
            else if (line.Contains(" M "))
            {
                Unit.TryGetValue("M", out unitValue);
            }
            else if (line.Contains(" G "))
            {
                Unit.TryGetValue("G", out unitValue);
            }
            else if (line.Contains(" T "))
            {
                Unit.TryGetValue("T", out unitValue);
            }

            double.TryParse(spt[step - 1], out double valueLow);
            double.TryParse(spt[step + 1], out double valueHigh);
            limitlow = valueLow * unitValue;
            limithigh = valueHigh * unitValue;
        }

    }
    public class LimitRow
    {
        public string Number = "";
        public int Site;
        public string TestName = "";
        public string Pin = "";
        public double Channel;
        public double LowLimit;
        public double Measured;
        public double HighLimit;
        public int Force = -1;
        public int Loc = -1;
    }

}
