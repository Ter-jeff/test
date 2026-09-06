using System.Diagnostics;
using System.Text.RegularExpressions;

namespace CommonLib.Datalog
{
    [DebuggerDisplay("{LineNo} : {Line}")]
    public partial class LineBase
    {
        [GeneratedRegex(@"\[Site\s*(?<site>\d+)\]", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexSite();
        [GeneratedRegex(@"\d+\.[a-z]\d+\w*", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexChannel();
        [GeneratedRegex(@"\d+\.\d+\w*", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        private static partial Regex MyRegexPowerBin();

        public static readonly Regex RegexSite = MyRegexSite();
        //eg.  23.x211h
        public static readonly Regex RegexChannel = MyRegexChannel();
        public static readonly Regex RegexPowerBin = MyRegexPowerBin();

        public int LineNo;
        public string Line = "";

        public static int GetChannelIndex(string[] spt)
        {
            return FindFirstMatchIndex(spt, RegexChannel);
        }

        public static int GetPowerBinningIndex(string[] spt)
        {
            return FindFirstMatchIndex(spt, RegexPowerBin);
        }

        private static int FindFirstMatchIndex(string[] spt, Regex regex)
        {
            for (int i = 0; i < spt.Length; i++)
            {
                if (regex.IsMatch(spt[i]))
                {
                    return i;
                }
            }
            return 0;
        }

        public static int GetMeasureIndex(int channelIndex, string[] spt)
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
