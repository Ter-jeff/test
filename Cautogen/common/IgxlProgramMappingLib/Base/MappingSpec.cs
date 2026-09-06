using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Cautogen.common.IgxlProgramMappingLib.Base
{
    public class MappingSpec
    {
        public HashSet<string> AcCategory = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> Timeset = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> DcCategoryLevel = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> PatternSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public string FastestAcCategory
        {
            get
            {
                return AcCategory.OrderByDescending((x) =>
                {
                    Match freqRegMatch = Regex.Match(x, @"_(?<shiftFreq>\d+)Mhz", RegexOptions.IgnoreCase);
                    if (freqRegMatch.Success)
                    {
                        if (int.TryParse(freqRegMatch.Groups["shiftFreq"].ToString(), out int parseOut))
                        {
                            return parseOut;
                        }

                        return 0;
                    }
                    else
                    {
                        return 0;
                    }
                }).FirstOrDefault() ?? "";
            }
        }

        public void Add(MappingSpec append)
        {
            append.AcCategory.ToList().ForEach(x => AcCategory.Add(x));
            append.Timeset.ToList().ForEach(x => Timeset.Add(x));
            append.DcCategoryLevel.ToList().ForEach(x => DcCategoryLevel.Add(x));
            append.PatternSet.ToList().ForEach(x => PatternSet.Add(x));
        }
    }
}
