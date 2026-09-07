using System.Text.RegularExpressions;

namespace DebugPlanReaderLib.DebugPlan
{
    public class PatternDate
    {
        public PatternDate(string pattern, string index = "", string subIndex = "")
        {
            OriName = pattern;
            Index = index;
            SubIndex = subIndex;
        }

        public string Name
        {
            get
            {
                if (WithDate)
                {
                    return OriName + "_PAT";
                }

                return OriName;
            }
            set { OriName = value; }
        }

        public bool WithDate
        {
            get { return Regex.IsMatch(OriName, @"\d+_\w\d+_\d+$"); }
        }

        public string OriName { get; set; }
        public string Index { get; set; }
        public string SubIndex { get; set; }
        public string Version { get; set; } = "";
        public bool SelsramDigSrc { get; set; } = false;
        public string DigSrcBitSize { get; set; }
        public string DigSrcPin { get; set; }
        public string DigSrcEQ { get; set; }
        public string DigSrcSeg { get; set; }
        public string DigSrcBits { get; set; }
        public string DigSrcBitSizeWithSubIndex
        {
            get
            {
                if (string.IsNullOrEmpty(DigSrcBitSize))
                    return "";
                else
                {
                    return SubIndex + ":" + Regex.Replace(DigSrcBitSize, @"^\w+:", "");
                }
            }
        }
        public string DigSrcPinWithSubIndex
        {
            get
            {
                if (string.IsNullOrEmpty(DigSrcPin))
                    return "";
                else
                {
                    return SubIndex + ":" + Regex.Replace(DigSrcPin, @"^\w+:", "");
                }
            }
        }
        public string DigSrcEQWithSubIndex
        {
            get
            {
                if (string.IsNullOrEmpty(DigSrcEQ))
                    return "";
                else
                {
                    return SubIndex + ":" + Regex.Replace(DigSrcEQ, @"^\w+:", "");
                }
            }
        }
        public string DigSrcSegWithSubIndex
        {
            get
            {
                if (string.IsNullOrEmpty(DigSrcSeg))
                    return "";
                else
                {
                    return SubIndex + ":" + Regex.Replace(DigSrcSeg, @"^\w+:", "");
                }
            }
        }
    }
}
