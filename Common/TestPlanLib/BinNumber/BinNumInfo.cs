using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

namespace TestPlanLib.BinNumber
{
    public class BinNumInfo
    {
        public string Module { get; set; } = string.Empty;
        public string Category1 { get; set; } = string.Empty;
        public string Category2 { get; set; } = string.Empty;
        public readonly List<BinNumDef> BinNumDef = [];
        public int HardBin
        {
            get
            {
                return BinNumDef.FirstOrDefault()?.HardBin ?? 999;
            }
        }
        public string Status
        {
            get
            {
                if (BinNumDef.Count != 0)
                {
                    if (BinNumDef.FirstOrDefault()!.Result.Trim().EqualsIgnoreCase("B"))
                    {
                    }
                    else if (BinNumDef.FirstOrDefault()!.Result.Trim().EqualsIgnoreCase("G"))
                    {
                        return "Pass";
                    }
                    else if (BinNumDef.FirstOrDefault()!.Result.Trim().EqualsIgnoreCase("O"))
                    {
                        return "Fail-Stop";
                    }
                }
                return "Fail";
            }
        }
        public SortedSet<int> SoftBinNums
        {
            get
            {
                var result = new SortedSet<int>();
                foreach (BinNumDef def in BinNumDef)
                {
                    int fisrtNum = Math.Min(def.BinFirst, def.BinLast);
                    int lastNum = Math.Max(def.BinFirst, def.BinLast);
                    result.AddRange(Enumerable.Range(fisrtNum, lastNum - fisrtNum + 1));
                }
                return result;
            }
        }
        public override bool Equals(object? obj)
        {
            if (obj is BinNumInfo other)
            {
                return Module.EqualsIgnoreCase(other.Module) &&
                    Category1.EqualsIgnoreCase(other.Category1) &&
                    Category2.EqualsIgnoreCase(other.Category2);
            }
            return false;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = (hash * 31) + (Module?.ToLower().GetHashCode() ?? 0);
                hash = (hash * 31) + (Category1?.ToLower().GetHashCode() ?? 0);
                hash = (hash * 31) + (Category2?.ToLower().GetHashCode() ?? 0);
                return hash;
            }
        }
    }
}
