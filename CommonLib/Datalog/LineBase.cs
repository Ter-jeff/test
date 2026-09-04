using System;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace CommonLib.Datalog
{
    [Serializable]
    [DebuggerDisplay("{LineNo} : {Line}")]
    public class LineBase
    {
        public readonly static Regex RegexSite = new Regex(@"\[Site\s*(?<site>\d+)\]", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        //eg.  23.x211h
        public readonly static Regex RegxChannel = new Regex(@"\d+\.[a-z]\d+\w*", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        public readonly static Regex RegxPowerBin = new Regex(@"\d+\.\d+\w*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public int LineNo;
        public string Line = "";

        public int GetChannelIndex(string[] spt)
        {
            int channelIndex = 0;
            for (int i = 0; i < spt.Length; i++)
            {
                if (RegxChannel.IsMatch(spt[i]))
                {
                    channelIndex = i;
                    break;
                }
            }
            return channelIndex;
        }

        public int GetPowerBinningIndex(string[] spt)
        {
            int channelIndex = 0;
            for (int i = 0; i < spt.Length; i++)
            {
                if (RegxPowerBin.IsMatch(spt[i]))
                {
                    channelIndex = i;
                    break;
                }
            }
            return channelIndex;
        }

        public int GetMeasureIndex(int channelIndex, string[] spt)
        {
            int step = 0;
            int moreTwoStep = 0;
            for (int i = channelIndex + 1; i < spt.Length; i++)
            {
                if (double.TryParse(spt[i], out _) || spt[i] == "N/A")
                {
                    moreTwoStep++;
                    if (moreTwoStep == 2)
                    {
                        step = i;
                        break;
                    }
                }
            }
            return step;
        }

        public int GetSite()
        {
            if (int.TryParse(RegexSite.Match(Line).Groups["site"].ToString(), out int site))
            {
                return site;
            }

            return -1;
        }
    }
}
