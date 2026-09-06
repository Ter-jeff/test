using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.common.IgxlProgramMappingLib.Base;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.common.IgxlProgramMappingLib
{
    public class MappingResult : List<MappingSpec>
    {
        private List<InstanceSheet> _InstanceSheets;
        private List<PatSetSheet> _PatSetSheets;
        private MappingDict _MappingDict = new MappingDict();
        private readonly Regex _regContainPerformanceModeByPattern = new Regex(@"(?!Mbist)^(?<pmode>M([a-zA-Z]){1}([a-zA-Z0-9]){1}(?<modenumber>[a-fA-F0-9]{2,3}))", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public MappingResult(List<InstanceSheet> instanceSheets, List<PatSetSheet> patSetsSheets)
        {
            _InstanceSheets = instanceSheets;
            _PatSetSheets = patSetsSheets;            
            _CreateMappingDict();
        }

        private void _CreateMappingDict()
        {
            var patSetsDict = _PatSetSheets.SelectMany(p => p.Rows).Where(q => !string.IsNullOrEmpty(q.PatSetName))
                                .GroupBy(p => p.PatSetName.ToUpper()).ToDictionary(p => p.Key.ToUpper(), p => p.ToList());
            foreach (InstanceSheet instanceSheet in _InstanceSheets)
            {
                foreach (InstanceRow instance in instanceSheet.Rows)
                {
                    if (instance.TestName.StartsWith("HFL_", StringComparison.OrdinalIgnoreCase)
                        || instance.TestName.StartsWith("HFH_", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (instance.Args.Count == 0)
                    {
                        continue;
                    }

                    string payload = instance.Args[0].ToUpper();

                    if (!patSetsDict.ContainsKey(payload))
                    {
                        var argList = instance.ArgList.Split(',').Select(x=>x.ToUpper()).ToList();
                        int payload1ArgIndex = argList.IndexOf("PayLoad_Patt1".ToUpper());

                        payload1ArgIndex = payload1ArgIndex == -1 ? argList.IndexOf("PATTERNS".ToUpper()) : payload1ArgIndex;

                        if (payload1ArgIndex != -1)
                        {
                            payload = instance.Args[payload1ArgIndex].ToUpper();
                            if (!patSetsDict.ContainsKey(payload))
                            {
                                continue;
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                        
                    foreach (PatSet patSet in patSetsDict[payload])
                    {
                        List<MappingKey> mappingKeys = _CreateMapping(patSet);
                        var mappingSpec = new MappingSpec();
                        mappingSpec.AcCategory.Add(instance.AcCategory);
                        mappingSpec.Timeset.Add(instance.TimeSets);
                        mappingSpec.DcCategoryLevel.Add(string.Format("{0}:{1}", instance.DcCategory, instance.PinLevels));
                        mappingSpec.PatternSet.Add(payload);
                        foreach (MappingKey mappingKey in mappingKeys)
                        {
                            _MappingDict.Add(mappingKey, mappingSpec);
                        }
                    }                    
                }
            } 
        }

        private List<MappingKey> _CreateMapping(PatSet patSet)
        {
            string pmode = "";
            var mappingKeys = new List<MappingKey>();
            var allPatterns = new List<string>();
            foreach (PatSetRow patRow in patSet.PatSetRows)
            {
                if (string.IsNullOrEmpty(patRow.File))
                {
                    continue;
                }
                string patName = (patRow.File.Contains("\\") || patRow.File.Contains("/")) ? patSet.PatSetName : patRow.File;
                string[] patSyntax = patName.Split('_');
                if (patSyntax.Length >= 10 && string.IsNullOrEmpty(pmode))
                {
                    if (_regContainPerformanceModeByPattern.IsMatch(patSyntax[9]) && !patSyntax[9].EndsWith("000"))
                    {
                        pmode = patSyntax[9];
                    }
                }

                if (patSyntax.Length >= 4)
                {
                    if (patSyntax[3].StartsWith("PL", StringComparison.OrdinalIgnoreCase) || patSyntax[3].Equals("FULP", StringComparison.OrdinalIgnoreCase))
                    {
                        MappingKey mappingKey = new MappingKey();
                        mappingKey.PatternSet = patSet.PatSetName;
                        mappingKey.Payload = patName;
                        mappingKey.Pmode = pmode;
                        mappingKeys.Add(mappingKey);
                    }
                }
                allPatterns.Add(patName.ToUpper());
            }
            mappingKeys.ForEach(x => x.AllPatterns = allPatterns);
            return mappingKeys;
        }
        public MappingKey CreateMappingCondition(List<string> patterns, string planPayload1)
        {
            var mappingKey = new MappingKey();
            foreach(string pattern in patterns)
            {
                string[] patSyntax = pattern.ToUpper().Split('_');

                if (patSyntax.Length >= 10 && string.IsNullOrEmpty(mappingKey.Pmode))
                {
                    if (_regContainPerformanceModeByPattern.IsMatch(patSyntax[9]) && !patSyntax[9].EndsWith("000"))
                    {
                        mappingKey.Pmode = patSyntax[9];
                    }
                }
                if (patSyntax.Length > 4)
                {
                    if (string.IsNullOrEmpty(mappingKey.Payload) && (patSyntax[3].StartsWith("PL", StringComparison.OrdinalIgnoreCase) || patSyntax[3].Equals("FULP", StringComparison.OrdinalIgnoreCase)))
                    {
                        mappingKey.Payload = pattern.ToUpper();
                    }
                }
            }
            if (!string.IsNullOrEmpty(planPayload1))
            {
                mappingKey.Payload = planPayload1;
            }
            mappingKey.AllPatterns = patterns.Select(x=> x.ToUpper()).ToList();
            return mappingKey;
        }
        public MappingSpec GenMappingIntoInstance(MappingKey mappingCondition, bool isHip)
        {                       
            if (!isHip)
            {
                return  _MappingDict.Query(mappingCondition);
            }
            else
            {
                return _MappingDict.QueryHip(mappingCondition);
            }
        }
    }
}
