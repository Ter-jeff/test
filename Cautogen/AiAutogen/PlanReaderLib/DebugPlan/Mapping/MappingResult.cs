using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using DebugPlanReaderLib.DebugPlan.Mapping.Base;

using IgxlLib.IgxlBase;

using IgxlLib.IgxlSheets;

namespace DebugPlanReaderLib.DebugPlan.Mapping
{
    public class MappingResult : List<MappingSpec>
    {
        private List<InstanceSheet> _InstanceSheets;
        private List<PatSetSheet> _PatSetSheets;
        private List<AiTestPlanSheet> _AiTestPlanSheets;
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
                    if (instance.Args.Count == 0)
                    {
                        continue;
                    }

                    string payload = instance.Args[0].ToUpper();

                    if (!patSetsDict.ContainsKey(payload))
                    {
                        var argList = instance.ArgList.Split(',').Select(x => x.ToUpper()).ToList();
                        int payload1ArgIndex = argList.IndexOf("PayLoad_Patt1".ToUpper());
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
                        mappingSpec.DcLevels.Add(string.Format("{0};{1}", instance.DcCategory, instance.PinLevels));
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
            foreach (PatSetRow patRow in patSet.PatSetRows)
            {
                string[] patSyntax = patRow.File.Contains("\\") ? patRow.PatternSet.Split('_') : patRow.File.Split('_');

                if (patSyntax.Length >= 10 && string.IsNullOrEmpty(pmode))
                {
                    if (_regContainPerformanceModeByPattern.IsMatch(patSyntax[9]) && !patSyntax[9].EndsWith("000"))
                    {
                        pmode = patSyntax[9];
                    }
                }

                if (patSyntax.Length >= 4)
                {
                    if (patSyntax[3].StartsWith("PL", StringComparison.OrdinalIgnoreCase))
                    {
                        MappingKey mappingKey = new MappingKey();
                        mappingKey.PatternSet = patRow.PatternSet;
                        mappingKey.Payload = patRow.File.Contains("\\") ? patRow.PatternSet : patRow.File;
                        mappingKey.Pmode = pmode;
                        mappingKeys.Add(mappingKey);
                    }
                }
            }
            return mappingKeys;
        }

        private MappingKey _CreateMappingCondition(AiTestPlanRow row)
        {
            var mappingKey = new MappingKey();
            foreach (PatternDate init in row.Inits)
            {
                string[] patSyntax = init.OriName.ToUpper().Split('_');

                if (patSyntax.Length >= 10 && string.IsNullOrEmpty(mappingKey.Pmode))
                {
                    if (_regContainPerformanceModeByPattern.IsMatch(patSyntax[9]) && !patSyntax[9].EndsWith("000"))
                    {
                        mappingKey.Pmode = patSyntax[9];
                    }
                }
            }

            if (!string.IsNullOrEmpty(row.MappingPattern))
            {
                mappingKey.Payload = row.MappingPattern.ToUpper();
            }
            return mappingKey;
        }

        public void GenMappingIntoInstance(List<AiTestPlanSheet> planSheets)
        {
            foreach (AiTestPlanSheet sheet in planSheets)
            {
                foreach (AiTestPlanRow row in sheet.Rows)
                {
                    var mappingCondition = _CreateMappingCondition(row);
                    var mappingSpec = _MappingDict.Query(mappingCondition);
                    if (mappingSpec != null)
                    {
                        row.AcCategoryMapping = mappingSpec.AcCategory.ToList();
                        row.TimesetMapping = mappingSpec.Timeset.ToList();
                        row.DcLevelsMapping = mappingSpec.DcLevels.ToList();
                    }
                }
            }
        }
    }
}
