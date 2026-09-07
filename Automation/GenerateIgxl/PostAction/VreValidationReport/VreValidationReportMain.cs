using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Automation.Static;

using CommonLib.Extension;

using CommonReaderLib.PatternListCsv;

using IgxlLib;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.VreValidationReport
{
    public class VreValidationReportMain
    {
        private Dictionary<string, HashSet<string>> _subflowJobInfo = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, SubFlowSheet> _allSubflows = new Dictionary<string, SubFlowSheet>(StringComparer.OrdinalIgnoreCase);
        private HashSet<InstanceInfo> _allInstanceInfos = new HashSet<InstanceInfo>();
        private List<MainFlow> _entryFlows;
        private HashSet<string> _bistHarvPatterns;
        private HashSet<string> _scanHarvPatterns;
        private Dictionary<string, string> _patternVerDic = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        internal static readonly Regex _regexTd = new Regex(@"_PL\w{2}_\w+_TDF_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        public VreValidationReportMain(HashSet<string> bistHarvPatterns, HashSet<string> scanHarvPatterns, List<string> subProgramName = null)
        {
            _bistHarvPatterns = bistHarvPatterns;
            _scanHarvPatterns = scanHarvPatterns;
            _entryFlows = TestProgram.IgxlWorkBk.MainFlowSheets.Values.ToList();
        }

        public void WorkFlow()
        {
            CreatePatternVersionDic();
            CollectSubflowJobInfo();
            CollectInstanceInfo();
            FillHarvInfo();
            AddActiveJobInfo();
            WriteReport();
        }
        private void CreatePatternVersionDic()
        {
            PatSetSheet patsetsAll = TestProgram.IgxlWorkBk.GetPatSetsSheet(IgxlWorkBook.PatSetsAll, FolderStructure.DirPatSetsAll);
            _patternVerDic = patsetsAll.PatSetRowDic.Values.ToDictionary(patset => patset.PatSetName, patset => Path.GetFileNameWithoutExtension(patset.PatSetRows.FirstOrDefault().File.Split(":").FirstOrDefault()));
        }
        private void CollectSubflowJobInfo()
        {
            _entryFlows = TestProgram.IgxlWorkBk.MainFlowSheets.Values.ToList();
            _allSubflows = TestProgram.IgxlWorkBk.SubFlowSheets.ToDictionary(x => x.Value.Name, x => x.Value);
            _subflowJobInfo = TestProgram.IgxlWorkBk.SubFlowSheets.ToDictionary(x => x.Value.Name, x => new HashSet<string>(StringComparer.OrdinalIgnoreCase));

            foreach (MainFlow entryFlow in _entryFlows)
            {
                _allSubflows.Add(entryFlow.Name, entryFlow);
                foreach (string currentJob in entryFlow.JobNames)
                {
                    foreach (FlowRow flowRow in entryFlow.Rows)
                    {
                        if (flowRow.Opcode != OpCode.Call)
                        {
                            continue;
                        }
                        FillJobInfo(flowRow, currentJob);
                    }
                    if (!_subflowJobInfo.ContainsKey(entryFlow.Name))
                    {
                        _subflowJobInfo.Add(entryFlow.Name, new HashSet<string> { currentJob });
                    }
                }
            }
            return;
        }

        private void CollectInstanceInfo()
        {
            var functionPatternIndexDic = TestProgram.VbtFunctionLib.VbtLib.ToDictionary(x => x.FullFunctionName, x => x.PatternDic.Keys.ToList());
            var allPatsetDic = TestProgram.IgxlWorkBk.PatSetSheets.Where(x => !x.Value.Name.Equals("PatSets_All") && !x.Value.Name.Equals("PatSets_DashBoard")).SelectMany(x => x.Value.Rows).DistinctBy(x => x.PatSetName).ToDictionary(x => x.PatSetName, x => x.PatSetRows.Select(y => y.File).ToList());
            var instance = TestProgram.IgxlWorkBk.InsSheets.SelectMany(x => x.Value.Rows).DistinctBy(x => x.TestName).ToList();
            foreach (KeyValuePair<string, InstanceSheet> instanceSheet in TestProgram.IgxlWorkBk.InsSheets)
            {
                foreach (InstanceRow instanceRow in instanceSheet.Value.Rows)
                {
                    string block = instanceSheet.Value.SourceInfo.Block;
                    List<string> patternList = new List<string>();
                    if (functionPatternIndexDic.TryGetValue(instanceRow.VbtName, out List<string> argList))
                    {
                        foreach (string argName in argList)
                        {
                            string patternName = instanceRow.GetArgument(argName);
                            if (allPatsetDic.TryGetValue(patternName, out List<string> patterns))
                            {
                                patternList.AddRange(patterns);
                            }
                            else
                            {
                                patternList.Add(patternName);
                            }
                        }
                    }
                    if (block == nameof(EnumBlock.Scan))
                    {
                        block = ScanTypeJudgement(patternList);
                    }
                    _allInstanceInfos.Add(new InstanceInfo { InstanceRow = instanceRow, Patterns = patternList, Block = block });
                }
            }
        }
        private string ScanTypeJudgement(List<string> patternList) => patternList.Any(_regexTd.IsMatch) ? "Td" : nameof(EnumBlock.Scan);

        private void AddActiveJobInfo()
        {
            foreach (InstanceInfo instanceInfo in _allInstanceInfos)
            {
                string instanceName = instanceInfo.InstanceName;
                foreach (KeyValuePair<string, HashSet<string>> avaliableFlow in _subflowJobInfo.Where(x => x.Value.Any()))
                {
                    HashSet<string> subFlowJob = avaliableFlow.Value;
                    SubFlowSheet subflow = _allSubflows[avaliableFlow.Key];
                    foreach (string currentJob in _entryFlows.SelectMany(x => x.JobNames).Distinct())
                    {
                        if (!subFlowJob.Contains(currentJob))
                        {
                            continue;
                        }

                        bool isActive = subflow.Rows.Exists(x => x.Opcode == OpCode.Test && x.Parameter.EqualsIgnoreCase(instanceName) && (string.IsNullOrEmpty(x.Job) || x.GetJobs().Exists(job => job.Equals(currentJob))));
                        if (isActive)
                        {
                            instanceInfo.ActiveJobs.Add(currentJob);
                        }
                    }
                }
            }
        }
        private void WriteReport()
        {
            IEnumerable<string> allJobs = _subflowJobInfo.Values.SelectMany(x => x).Distinct();
            var totalBlocks = new List<string> { nameof(EnumBlock.Scan), "Td", nameof(EnumBlock.Mbist) };
            foreach (string job in allJobs)
            {
                var instanceInfoByJob = _allInstanceInfos.Where(x => x.ActiveJobs.Contains(job)).ToList();
                foreach (string block in totalBlocks)
                {
                    IEnumerable<InstanceInfo> allInstanceInfoByBlock = instanceInfoByJob.Where(x => x.Block.EqualsIgnoreCase(block));
                    IEnumerable<InstanceInfo> harvInstance = allInstanceInfoByBlock.Where(x => x.IsHarv);
                    IEnumerable<InstanceInfo> nonHarvInst = allInstanceInfoByBlock.Where(x => !x.IsHarv);
                    WirteDetail(block, true, job, harvInstance);
                    WirteDetail(block, false, job, nonHarvInst);
                }
            }
        }

        private void WirteDetail(string block, bool isHarv, string job, IEnumerable<InstanceInfo> instances)
        {
            string harvType = isHarv ? "Harvesting" : "non_Harvesting";
            string outputPath = Path.Combine(FolderStructure.DirVreValid, $"{job}_{block}_{harvType}.csv");
            if (!instances.Any())
            {
                return;
            }
            using (var detailWriter = new StreamWriter(outputPath, false, new UTF8Encoding(true)))
            {
                List<string> titles = new List<string> { "Instance" };
                int maxPatternCount = instances.Max(instance => instance.Patterns.Count());
                for (int number = 1; number <= maxPatternCount; number++)
                {
                    titles.Add($"Pattern{number}");
                }
                detailWriter.WriteLine(string.Join(",", titles));
                foreach (InstanceInfo instance in instances)
                {
                    List<string> contentList = new List<string> { instance.InstanceName };
                    List<string> patternList = instance.Patterns;
                    patternList.ForEach(pattern => contentList.Add(_patternVerDic.ContainsKey(pattern) ? _patternVerDic[pattern] : pattern));
                    detailWriter.WriteLine(string.Join(",", contentList));
                }
            }
        }
        private void FillJobInfo(FlowRow flowRow, string currentJob)
        {
            string subflowName = flowRow.Parameter;
            if (!_subflowJobInfo.ContainsKey(subflowName))
            {
                return;
            }
            List<string> activeJobs = flowRow.GetJobs();
            if (activeJobs.Any())
            {
                if (!activeJobs.Exists(x => x.EqualsIgnoreCase(currentJob)))
                {
                    return;
                }
                if (activeJobs.Exists(x => x.EqualsIgnoreCase("!" + currentJob)))
                {
                    return;
                }
            }
            if (!_subflowJobInfo[subflowName].Contains(currentJob))
            {
                _subflowJobInfo[subflowName].Add(currentJob);
            }
            if (_allSubflows.TryGetValue(subflowName, out SubFlowSheet targetFlow))
            {
                foreach (FlowRow row in targetFlow.Rows)
                {
                    if (row.Opcode != OpCode.Call)
                    {
                        continue;
                    }
                    FillJobInfo(row, currentJob);
                }
            }
        }
        private void FillHarvInfo()
        {
            foreach (InstanceInfo instanceInfo in _allInstanceInfos)
            {
                switch (instanceInfo.Block)
                {
                    case nameof(EnumBlock.Scan):
                        {
                            instanceInfo.IsHarv = instanceInfo.InstanceRow.GetArgument("isHarvesting").EqualsIgnoreCase("TRUE");
                            if (instanceInfo.IsHarv)
                            {
                                instanceInfo.HarvPatterns.AddRange(instanceInfo.Patterns.Where(_scanHarvPatterns.Contains));
                            }
                            break;
                        }
                    case "Td":
                        {
                            instanceInfo.IsHarv = instanceInfo.InstanceRow.GetArgument("isHarvesting").EqualsIgnoreCase("TRUE");
                            if (instanceInfo.IsHarv)
                            {
                                instanceInfo.HarvPatterns.AddRange(instanceInfo.Patterns.Where(_scanHarvPatterns.Contains));
                            }
                            break;
                        }
                    case nameof(EnumBlock.Harvest):
                        {
                            instanceInfo.IsHarv = true;
                            break;
                        }
                    case nameof(EnumBlock.Mbist):
                        {
                            List<string> harvPatterns = instanceInfo.Patterns.FindAll(_bistHarvPatterns.Contains);
                            instanceInfo.IsHarv = harvPatterns.Any();
                            if (instanceInfo.IsHarv)
                            {
                                instanceInfo.HarvPatterns.AddRange(harvPatterns);
                            }
                            break;
                        }
                }
            }
        }
    }
    public class InstanceInfo
    {
        public string InstanceName
        {
            get
            {
                return InstanceRow.TestName;
            }
        }
        public InstanceRow InstanceRow;
        public List<string> Patterns = new();
        public bool IsHarv = false;
        public HashSet<string> ActiveJobs = new();
        public string Block = "";
        public List<string> HarvPatterns = new();
    }
}
