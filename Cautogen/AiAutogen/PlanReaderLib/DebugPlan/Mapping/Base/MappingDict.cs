using System;
using System.Collections.Generic;
using System.Linq;

namespace DebugPlanReaderLib.DebugPlan.Mapping.Base
{
    public class MappingDict : Dictionary<MappingKey, MappingSpec>
    {
        public new void Add(MappingKey key, MappingSpec value)
        {
            if (this.ContainsKey(key))
            {
                this[key].Add(value);
            }
            else
            {
                this[key] = value;
            }
        }

        public MappingSpec Query(MappingKey key)
        {
            MappingSpec result = null;
            string payload = key.Payload;
            string pmode = key.Pmode;

            if (!string.IsNullOrEmpty(key.Payload))
            {
                var mapping = this.Where(x => x.Key.Payload == key.Payload).ToDictionary(x => x.Key, x => x.Value);
                if (mapping != null && mapping.Select(x => x.Value.AcCategory).Count() > 1)
                {
                    if (!string.IsNullOrEmpty(key.Pmode))
                    {
                        var pmodeMapping = mapping.Where(x => x.Key.Pmode.Equals(key.Pmode, StringComparison.OrdinalIgnoreCase)).ToDictionary(x => x.Key, x => x.Value);
                        if (pmodeMapping.Select(x => x.Value.AcCategory).Count() > 0)
                        {
                            mapping = pmodeMapping;
                        }
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(key.Domain)
                        && !string.IsNullOrEmpty(key.Cluster)
                        && !string.IsNullOrEmpty(key.BlockType))
                    {
                        mapping = this.Where(x => x.Key.Domain == key.Domain
                                                    && x.Key.Cluster == key.Cluster
                                                    && x.Key.BlockType == key.BlockType)
                                            .ToDictionary(x => x.Key, x => x.Value);
                    }
                    if (mapping != null && mapping.Select(x => x.Value.AcCategory).Count() > 1)
                    {
                        if (!string.IsNullOrEmpty(key.Pmode))
                        {
                            var pmodeMapping = mapping.Where(x => x.Key.Pmode.Equals(key.Pmode, StringComparison.OrdinalIgnoreCase)).ToDictionary(x => x.Key, x => x.Value);
                            if (pmodeMapping.Select(x => x.Value.AcCategory).Count() > 0)
                            {
                                mapping = pmodeMapping;
                            }
                        }
                    }
                }
                if (mapping != null)
                {
                    result = new MappingSpec();
                    result.AcCategory = new HashSet<string>(mapping.Values.SelectMany(x => x.AcCategory), StringComparer.OrdinalIgnoreCase);
                    result.Timeset = new HashSet<string>(mapping.Values.SelectMany(x => x.Timeset), StringComparer.OrdinalIgnoreCase);
                    result.DcLevels = new HashSet<string>(mapping.Values.SelectMany(x => x.DcLevels), StringComparer.OrdinalIgnoreCase);
                }
            }
            return result;
        }
    }
}
