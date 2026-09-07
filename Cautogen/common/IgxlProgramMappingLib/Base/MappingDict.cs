using System;
using System.Collections.Generic;
using System.Linq;

namespace Cautogen.common.IgxlProgramMappingLib.Base
{
    public class MappingDict : Dictionary<MappingKey, MappingSpec>
    {
        public new void Add(MappingKey key, MappingSpec value)
        {
            if (ContainsKey(key))
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
            MappingSpec result = new MappingSpec();
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
                        if (pmodeMapping.Select(x => x.Value.AcCategory).Any())
                        {
                            mapping = pmodeMapping;
                        }
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(key.Domain)
                        && !string.IsNullOrEmpty(key.Cluster)
                        && !string.IsNullOrEmpty(key.BlockType)
                        && !string.IsNullOrEmpty(key.SiDmMode))
                    {
                        mapping = this.Where(x => x.Key.Domain == key.Domain
                                                    && x.Key.Cluster == key.Cluster
                                                    && x.Key.BlockType == key.BlockType
                                                    && x.Key.SiDmMode == key.SiDmMode)
                                            .ToDictionary(x => x.Key, x => x.Value);
                    }

                    if (mapping != null && mapping.Select(x => x.Value.AcCategory).Count() > 1)
                    {
                        if (!string.IsNullOrEmpty(key.Pmode))
                        {
                            var pmodeMapping = mapping.Where(x => x.Key.Pmode.Equals(key.Pmode, StringComparison.OrdinalIgnoreCase)).ToDictionary(x => x.Key, x => x.Value);
                            if (pmodeMapping.Select(x => x.Value.AcCategory).Any())
                            {
                                mapping = pmodeMapping;
                            }
                        }
                    }
                }
                if (mapping != null)
                {
                    result.AcCategory = new HashSet<string>(mapping.Values.SelectMany(x => x.AcCategory), StringComparer.OrdinalIgnoreCase);
                    result.Timeset = new HashSet<string>(mapping.Values.SelectMany(x => x.Timeset), StringComparer.OrdinalIgnoreCase);
                    result.DcCategoryLevel = new HashSet<string>(mapping.Values.SelectMany(x => x.DcCategoryLevel), StringComparer.OrdinalIgnoreCase);
                    result.PatternSet = new HashSet<string>(mapping.Values.SelectMany(x => x.PatternSet), StringComparer.OrdinalIgnoreCase);
                }
            }
            return result;
        }
        public MappingSpec QueryHip(MappingKey key)
        {
            MappingSpec result = new MappingSpec();
            string payload = key.Payload;
            List<string> allPatterns = key.AllPatterns;

            if (!string.IsNullOrEmpty(key.Payload))
            {
                var mapping = this.Where(x => x.Key.Payload == key.Payload
                    && x.Key.AllPatterns.SequenceEqual(allPatterns)).ToDictionary(x => x.Key, x => x.Value);

                if (mapping.Count == 0 || mapping == null)
                {
                    //mapping for payload only
                    mapping = this.Where(x => x.Key.Payload == key.Payload).ToDictionary(x => x.Key, x => x.Value);
                }

                if (mapping != null)
                {
                    result.AcCategory = new HashSet<string>(mapping.Values.SelectMany(x => x.AcCategory), StringComparer.OrdinalIgnoreCase);
                    result.Timeset = new HashSet<string>(mapping.Values.SelectMany(x => x.Timeset), StringComparer.OrdinalIgnoreCase);
                    result.DcCategoryLevel = new HashSet<string>(mapping.Values.SelectMany(x => x.DcCategoryLevel), StringComparer.OrdinalIgnoreCase);
                    result.PatternSet = new HashSet<string>(mapping.Values.SelectMany(x => x.PatternSet), StringComparer.OrdinalIgnoreCase);
                }
            }
            return result;
        }
    }
}
