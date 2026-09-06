using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace EfuseCheckCmdLib.Base
{
    public partial class OneTouchDown(string datafile)
    {
        public static readonly Regex RegXy = XyRegex();
        public static readonly Regex RegIsNumeric = IsNumericRegex();

        [GeneratedRegex(@"Site\s+X_Coord\s+Y_Coord", RegexOptions.IgnoreCase)]
        private static partial Regex XyRegex();

        [GeneratedRegex(@"^\s*\d+(\s+\d+)*\s*$")]
        private static partial Regex IsNumericRegex();

        [GeneratedRegex(@"\s+")]
        private static partial Regex WhitespaceRegex();

        private readonly string _fileName = datafile;

        public List<List<string>> SplitDataFile()
        {
            bool foundXy = false;
            List<List<string>> segmentData = [];
            List<string> curSegmentData = [];
            List<string> lines = [.. File.ReadAllLines(_fileName)];

            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];
                if (foundXy)
                {
                    if (line.Contains('-'))
                    {
                        curSegmentData.Add(line);
                        continue;
                    }
                    string[] dieTest = WhitespaceRegex().Split(line.Trim());
                    if (dieTest.Length == 3 && RegIsNumeric.IsMatch(line))
                    {
                        curSegmentData.Add(line);
                        continue;
                    }
                    else
                    {
                        segmentData.Add([.. curSegmentData]);
                        curSegmentData.Clear();
                        foundXy = false;
                    }
                }

                if (RegXy.IsMatch(line))
                {
                    foundXy = true;
                }
                curSegmentData.Add(line);
            }

            if (curSegmentData.Count > 0)
            {
                segmentData.Add([.. curSegmentData]);
            }
            return segmentData;
        }
    }
}
