using System;
using System.Collections.Generic;
using System.Linq;

namespace DebugPlanReaderLib.DebugPlan.Mapping.Base
{
    public class MappingSpec
    {
        public HashSet<string> AcCategory = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> Timeset = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> DcLevels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public void Add(MappingSpec append)
        {
            append.AcCategory.ToList().ForEach(x => this.AcCategory.Add(x));
            append.Timeset.ToList().ForEach(x => this.Timeset.Add(x));
            append.DcLevels.ToList().ForEach(x => this.DcLevels.Add(x));
        }
    }
}
