using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

namespace Cautogen.AutoCZ.CharPostProcessor.Controller
{
    public class ProdProg
    {
        /* Field */
        private static int _jobIndex;
        public static ChannelMapSheet ChannelMapSheet;
        public static Dictionary<string, List<string>> AllTimeSetUseAcDictionary = new Dictionary<string, List<string>>();
        public static List<InstanceRow> AllTestInstances = new List<InstanceRow>();
        public static Dictionary<string, InstanceRow> AllTestInstancesDict = new Dictionary<string, InstanceRow>(StringComparer.OrdinalIgnoreCase);
        public static Dictionary<string, List<InstanceRow>> SamePatternInstanceDict = new Dictionary<string, List<InstanceRow>>();
        public static List<KeyValuePair<string, List<string>>> AllPatSetRows = new List<KeyValuePair<string, List<string>>>();
        public static SubFlowSheet TmpsFlow = null;
        public static SubFlowSheet NwireFlow = null;
        public static TimeSetBasicSheet NwireTimeset = null;

        /* Member function */
        public static void Parse(string jobName, string channelMapSheetName)
        {
            Response.Report("Parasing test program... ");

            // clear previous result
            AllTestInstances = new List<InstanceRow>();
            SamePatternInstanceDict = new Dictionary<string, List<InstanceRow>>();

            try
            {
                _jobIndex = LocalSpecs.TestProgram.GetJobList().FindIndex(a => a.Equals(jobName, StringComparison.OrdinalIgnoreCase));

                if (_jobIndex == -1)
                {
                    throw new Exception("Can not find job " + jobName + " in the prod program.");
                }

                _UpdateSheets(channelMapSheetName);

                // update inst, flowSteps and pinDict
                AllTestInstancesDict = _UpdateInstList();
                _UpdateFlowSteps(AllTestInstancesDict);
                _UpdatePinDic();
                _UpdatePinTypeInChannelDic();

                // update PatternInstanceDict
                List<InstanceRow> arg0Instance = AllTestInstances.FindAll(p => p.Args.Count == 0);
                foreach (InstanceRow testInstance in arg0Instance)
                {
                    AllTestInstances.RemoveAll(p => p.Equals(testInstance));
                }

                SamePatternInstanceDict = AllTestInstances.GroupBy(o => o.Args[0].ToUpper()).ToDictionary(g => g.Key, g => g.ToList());
                AllTimeSetUseAcDictionary = AllTestInstances
                    .Where(p => !string.IsNullOrEmpty(p.TimeSets))
                    .GroupBy(p => p.TimeSets.ToUpper())
                    .ToDictionary(p => p.Key, p => p.Select(k => k.AcCategory).ToList());

                AllPatSetRows = LocalSpecs.TestProgram.PatSetSheets
                    .SelectMany(x => x.Rows)
                    .Where(x => !string.IsNullOrEmpty(x.PatSetName))
                    .GroupBy(x => x.PatSetName)
                    .Select(g => new KeyValuePair<string, List<string>>(
                        g.Key,
                        g.SelectMany(y => y.PatSetRows)
                         .Select(z => z.File.ToUpper())
                         .Distinct()
                         .ToList()
                    ))
                    .ToList();

                TmpsFlow = LocalSpecs.TestProgram.FlowSheets.FirstOrDefault(p => Regex.IsMatch(p.Name, "Flow_TMPS", RegexOptions.IgnoreCase));
                NwireFlow = LocalSpecs.TestProgram.FlowSheets.FirstOrDefault(p => Regex.IsMatch(p.Name, "Flow_nWire_default", RegexOptions.IgnoreCase));
                NwireTimeset = LocalSpecs.TestProgram.TimeSetSheets.FirstOrDefault(p => Regex.IsMatch(p.Name, "Timeset_nwire", RegexOptions.IgnoreCase));
            }
            catch (Exception e)
            {
                throw new Exception("Parsing test program failed! " + e.Message);
            }
        }

        private static void _UpdateSheets(string channelMapSheetName)
        {
            // channel map
            if (string.IsNullOrEmpty(channelMapSheetName) && LocalSpecs.TestProgram.ChannelMapSheets.Any())
            {
                ChannelMapSheet = LocalSpecs.TestProgram.ChannelMapSheets.First();
            }
            else
            {
                ChannelMapSheet = LocalSpecs.TestProgram.ChannelMapSheets.
                    FirstOrDefault(p => p.Name.Equals(channelMapSheetName, StringComparison.OrdinalIgnoreCase));
            }

            // pin map
            //var pinMapSheetName = LocalSpecs.TestProgram.GetJobList()[_jobIndex].JobSheetFields.PinMapSheetName;


            // test instance

            //var instSheetNames = LocalSpecs.TestProgram.InstanceSheets.Select(p=>p.IgxlSheetName);

            //foreach (var instSheet in instSheetNames
            //    .Select(instSheetName => (TestInstanceSheet) LocalSpecs.TestProgram.GetSheetByName(instSheetName))
            //    .Where(instSheet => instSheet != null && instSheet.TestInstances.Count != 0))
            //{
            //    _testInstanceSheets.Add(instSheet);
            //}
        }

        private static void _UpdatePinDic()
        {
            /* Update LocalSpecs.ProgInfo.PinDic and PinGroupDic*/

            // push pin.name.Replace("_", "") -> pin.name to PinDic
            foreach (Pin pin in LocalSpecs.TestProgram.PinList)
            {
                string pinName = pin.PinName.Replace("_", "").ToUpper();
                if (!LocalSpecs.ProgInfo.PinDic.ContainsKey(pinName))
                {
                    LocalSpecs.ProgInfo.PinDic.Add(pinName, pin.PinName);
                }
            }

            // do the same for pin group

            foreach (PinGroup pinGroup in LocalSpecs.TestProgram.PinGroups)
            {
                string pinGroupName = pinGroup.PinName.Replace("_", "").ToUpper();
                if (LocalSpecs.ProgInfo.PinDic.ContainsKey(pinGroupName))
                {
                    continue;
                }

                LocalSpecs.ProgInfo.PinDic.Add(pinGroupName, pinGroup.PinName);
                LocalSpecs.ProgInfo.PinGroupDic.Add(pinGroupName, pinGroup);
            }
        }

        private static void _UpdatePinTypeInChannelDic()
        {
            foreach (ChannelMapRow channelRow in ChannelMapSheet.Rows)
            {
                string pinName = channelRow.DeviceUnderTestPinName.ToUpper();
                string pinInstrumentType = channelRow.Type.ToUpper();
                if (!LocalSpecs.ProgInfo.PinTypeInChannelDic.ContainsKey(pinName))
                {
                    LocalSpecs.ProgInfo.PinTypeInChannelDic.Add(pinName, pinInstrumentType);
                }
            }
        }

        private static void _UpdateFlowSteps(IDictionary<string, InstanceRow> instDict)
        {
            // containers
            var flowSteps = new Dictionary<string, FlowRow>();
            var userLimits = new Dictionary<string, Dictionary<string, List<FlowRow>>>(StringComparer.OrdinalIgnoreCase);  // pattern -> use-limit rows
            var firstUserLimitTestInstNameDict = new Dictionary<string, string>();  // pattern -> first test instance name

            // get main flow name
            SubFlowSheet mainFlowSheet = LocalSpecs.TestProgram.GetMainFlowSheet();
            if (mainFlowSheet == null)
            {
                return;
            }

            string mainFlowSheetName = LocalSpecs.TestProgram.GetMainFlowSheet().Name;
            if (mainFlowSheetName == "")
            {
                return;
            }

            // retrieve all the subFlowNames
            var subFlowNames = new List<string>();
            _UpdateSubFlowNames(mainFlowSheetName, subFlowNames);

            var skipOpcodeList = new List<string> { "characterize", "bintable" }; //, "test" };
            foreach (SubFlowSheet subFlow in LocalSpecs.TestProgram.FlowSheets
                .Where(p => subFlowNames.Exists(sheetName => sheetName.Equals(p.Name, StringComparison.OrdinalIgnoreCase))))
            {

                foreach (FlowRow flowRow in subFlow.Rows)
                {
                    flowRow.SheetName = subFlow.Name;
                    string instName = flowRow.Parameter.ToLower();

                    // handle use limit
                    if (flowRow.Opcode.ToLower() == "use-limit")
                    {
                        if (!instDict.TryGetValue(instName, out InstanceRow inst))
                        {
                            continue;
                        }

                        if (inst == null)
                        {
                            continue;
                        }

                        string patternName = "";
                        if (inst.Args.Count != 0)
                        {
                            patternName = inst.Args[0].ToLower();
                        }

                        string patternNameWithVoltage = patternName + _GetVoltageByInstName(instName);
                        string instanceSourceInfomation = flowRow.SheetName + ":" + instName;
                        if (!userLimits.ContainsKey(patternNameWithVoltage))
                        {
                            userLimits.Add(patternNameWithVoltage, new Dictionary<string, List<FlowRow>>(StringComparer.OrdinalIgnoreCase));
                            userLimits[patternNameWithVoltage][instanceSourceInfomation] = new List<FlowRow>();
                        }
                        else
                        {
                            if (!userLimits[patternNameWithVoltage].ContainsKey(instanceSourceInfomation))
                            {
                                userLimits[patternNameWithVoltage][instanceSourceInfomation] = new List<FlowRow>();
                            }
                        }
                        userLimits[patternNameWithVoltage][instanceSourceInfomation].Add(flowRow);
                        continue;
                    }

                    // skip if instName == ""
                    if (instName == "")
                    {
                        continue;
                    }

                    // skip if gated by ENV
                    if (!string.IsNullOrEmpty(flowRow.Env))
                    {
                        continue;
                    }

                    // skip if opCode in skipOpCodeList
                    if (skipOpcodeList.Contains(flowRow.Opcode.ToLower()))
                    {
                        continue;
                    }

                    // skip if flowRow is an instance and uses PP pattern.
                    if (instDict.TryGetValue(instName, out InstanceRow instance))
                    {
                        if (instance == null)
                        {
                            continue;
                        }
                        // if (Regex.IsMatch(instance.Args[0], @"^pp_", RegexOptions.IgnoreCase)) continue;
                    }

                    LocalSpecs.ProgInfo.AllFlowSteps.Add(flowRow);
                    if (!flowSteps.ContainsKey(instName))
                    {
                        flowSteps.Add(instName, flowRow);
                    }
                }
            }
            LocalSpecs.AllFlowStepsDic = flowSteps;
            LocalSpecs.UseLimitDict = userLimits;
        }

        private static Dictionary<string, InstanceRow> _UpdateInstList()
        {
            /* get non-duplicated test instance list from all test inst sheets records in job sheet */

            var instDict = new Dictionary<string, InstanceRow>(StringComparer.OrdinalIgnoreCase);
            //var instSheetNames = LocalSpecs.TestProgram.Jobs[_jobIndex].JobSheetFields.TestInstanceSheetNames;

            //foreach (var instance in from instSheetName in instSheetNames
            //    select (TestInstanceSheet) LocalSpecs.TestProgram.GetSheetByName(instSheetName)
            //    into instSheet
            //    where instSheet != null && instSheet.TestInstances.Count != 0
            //    from inst in instSheet.TestInstances
            //    where !instDict.ContainsKey(inst.TestName.ToLower())
            //    select inst)
            foreach (InstanceRow instance in LocalSpecs.TestProgram.InstanceSheets.SelectMany(p => p.Rows))
            {
                if (instance.TestName.StartsWith("HFL_", StringComparison.OrdinalIgnoreCase) ||
                    instance.TestName.StartsWith("HFH_", StringComparison.OrdinalIgnoreCase) ||
                    instance.TestName.StartsWith("HFLH_", StringComparison.OrdinalIgnoreCase) ||
                    instance.TestName.StartsWith("HFHL_", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                AllTestInstances.Add(instance);
                if (!instDict.ContainsKey(instance.TestName.ToLower()))
                {
                    instDict.Add(instance.TestName.ToLower(), instance);
                }
            }
            return instDict;
        }

        private static void _UpdateSubFlowNames(string flowSheetName, List<string> subFlows)
        {
            /* recursive collect all the subflow by the call tree rooted in main flow */
            SubFlowSheet flowSheet = LocalSpecs.TestProgram.FlowSheetsAll.
                FirstOrDefault(p => p.Name.Equals(flowSheetName, StringComparison.OrdinalIgnoreCase));
            if (flowSheet == null)
            {
                return;
            }

            foreach (FlowRow flowStep in flowSheet.Rows
                .Where(flowStep => flowStep.Opcode.ToLower() == "call")
                .Where(flowStep => subFlows.FindIndex(x => string.Equals(x, flowStep.Parameter, StringComparison.CurrentCultureIgnoreCase)) == -1))
            {
                subFlows.Add(flowStep.Parameter);
                _UpdateSubFlowNames(flowStep.Parameter, subFlows);
            }
        }

        private static string _GetVoltageByInstName(string instName)
        {
            if (Regex.IsMatch(instName, "(?<voltage>_NV|_HV|_LV)", RegexOptions.IgnoreCase))
            {
                Match match = Regex.Match(instName, "(?<voltage>_NV|_HV|_LV)", RegexOptions.IgnoreCase);
                return match.Groups["voltage"].ToString();
            }
            return "";
        }
    }
}
