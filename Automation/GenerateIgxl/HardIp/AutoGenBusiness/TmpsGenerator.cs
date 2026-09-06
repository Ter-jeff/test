using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.PostAction.Relay.RelayConst;
using Automation.Reader;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class TmpsGenerator
    {
        private const string InitialAdaptiveCoolingInstanceName = "InitialAdaptiveCooling";
        private const string ControlAdaptiveCoolingInstanceName = "ControlAdaptiveCooling";
        private const string CleanUpAdaptiveCoolingInstanceName = "CleanUpAdaptiveCooling";

        private readonly string _indexName;
        private readonly string _startIndex;
        private readonly string _endIndex;
        private readonly string _flagName;
        private readonly HashSet<string> _adaptiveCoolingItem;

        public TmpsGenerator(string indexName = "SrcCodeIndx", string startIndex = "1", string endIndex = "3", string flagName = "F_TMPS_Monitor")
        {
            _indexName = indexName;
            _startIndex = startIndex;
            _endIndex = endIndex;
            _flagName = flagName;
            _adaptiveCoolingItem = new ConfigData().AdaptiveCoolingItem;
        }

        public void GenTmps(Dictionary<string, SubFlowSheet> subFlowSheets)
        {
            var flowSheets = subFlowSheets.Select(x => x.Value).ToList();
            var tmpsSubflows = new List<SubFlowSheet>();
            if (_adaptiveCoolingItem.Any())
            {
                // Use block & subblock to find flow
                IEnumerable<string> tmpsItems = _adaptiveCoolingItem.Where(x => !x.StartsWith("IDS_", StringComparison.OrdinalIgnoreCase));
                IEnumerable<string> idsItems = _adaptiveCoolingItem.Where(x => x.StartsWith("IDS_", StringComparison.OrdinalIgnoreCase));

                var tmpsFlowRows = new List<FlowRow>();
                foreach (string item in tmpsItems)
                {
                    string block = item.Split('_')[0];
                    string subBlock = item.Split('_')[1];
                    SubFlowSheet flowSheet = flowSheets.FirstOrDefault(x => string.Equals(x.Name.ToLower().Replace("flow_hardip_", "").Replace("_", ""), block, StringComparison.CurrentCultureIgnoreCase));
                    if (flowSheet != null)
                    {
                        var flowRows = flowSheet.Rows.Where(x => x.Parameter.ToUpper().StartsWith($"{block}_{subBlock}_") && x.Parameter.ToUpper().EndsWith("_NV")).Select(x => x.Copy()).ToList();
                        foreach (FlowRow row in flowRows)
                        {
                            //hardcode change the name with MON
                            row.Parameter = row.Parameter.ToUpper().Replace("_" + subBlock, "_" + subBlock + "MON");
                        }
                        tmpsFlowRows.AddRange(flowRows);
                    }
                }

                var idsFlowRows = new List<FlowRow>();
                foreach (string item in idsItems)
                {
                    string block = item.Split('_')[0];
                    string subBlock = item.Split('_')[1];
                    SubFlowSheet idsFlowSheet = flowSheets.FirstOrDefault(x => x.Name.ContainsIgnoreCase("DCTEST_IDS"));
                    if (idsFlowSheet != null)
                    {
                        var idsRows = idsFlowSheet.Rows.Where(x => x.Opcode.Equals(OpCode.Test) && x.Parameter.ToUpper().StartsWith($"{block}_{subBlock}_") && x.Parameter.ToUpper().EndsWith("_NV")).Select(x => x.Copy()).ToList();
                        idsFlowRows.AddRange(idsRows);
                    }
                }
                if (tmpsFlowRows.Any())
                {
                    //for non-HIP usage
                    tmpsSubflows.Add(GenerateTmpsFlow(tmpsFlowRows, idsFlowRows, "Flow_TMPS"));
                    //for HIP usage
                    tmpsSubflows.Add(GenerateTmpsFlow(tmpsFlowRows, idsFlowRows, "Flow_TMPS_NO_RELAY", true));
                    //tmps instances
                    GenerateTmpsInstanceIntoCommon();
                }
            }
            foreach (SubFlowSheet tmpsSubflow in tmpsSubflows)
            {
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, tmpsSubflow);
            }
        }

        public SubFlowSheet GenerateTmpsFlow(List<FlowRow> flowRows, List<FlowRow> idsRows, string tmpsSheetName, bool noRelay = false)
        {
            string flowSheetName = tmpsSheetName;
            var flowSheet = new SubFlowSheet(flowSheetName, "HARDIP_TMPS");

            //Print start.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Print,
                Parameter = '"' + flowSheetName.Replace(" ", "_") + " Start " + '"'
            });
            //Initial instance.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Test,
                Parameter = InitialAdaptiveCoolingInstanceName
            });
            //Relay on.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = noRelay ? OpCode.Nop : OpCode.Test,
                Parameter = "Relay_On_HARDIP"
            });
            //When it ran with UFP, need to disable nWire.
            if (TestProgram.IgxlWorkBk.SubFlowSheets.Any(p => p.Value.Name.Equals($"{NwireSetting.ConFlownWire}Default_Disable", StringComparison.OrdinalIgnoreCase)))
            {
                IEnumerable<string> jobs = TestPlanStatic.JobInfoSheet?.Rows.Where(x => x.TesterType.ToUpper().Equals("UFP")).Select(x => x.JobName).ToList();
                flowSheet.AddRow(new FlowRow
                {
                    Job = jobs?.Any() == true ? string.Join(",", jobs) : "",
                    Opcode = jobs?.Any() == true ? OpCode.Call : OpCode.Nop,
                    Parameter = $"{NwireSetting.ConFlownWire}Default_Disable"
                });
            }
            //Assign Integer row.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.AssignInteger,
                Parameter = $"SrcCodeIndx1 {_endIndex}"
            });
            //Add new var-assign.
            if (!HardIpStatic.FlowUsedInteger.Contains("SrcCodeIndx1"))
            {
                HardIpStatic.FlowUsedInteger.Add("SrcCodeIndx1");
            }

            //Loop.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.For,
                Parameter = $"SrcCodeIndx={_startIndex};SrcCodeIndx<=SrcCodeIndx1;SrcCodeIndx++"
            });
            //Clear flag.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.FlagClear,
                Parameter = $"{_flagName}"
            });
            //Judge if it is the lastest loop.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.If,
                Parameter = "SrcCodeIndx!=SrcCodeIndx1"
            });
            //Normal spec.
            flowSheet.AddRows(GenerateTmpsBodyRows(flowRows, $"{_flagName}"));
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Else
            });
            //Relax spec.
            flowSheet.AddRows(GenerateTmpsBodyRows(flowRows, $"{_flagName}", true));
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.EndIf
            });
            //When adaptive cooling is disabled, just run one time.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Break,
                Enable = "!Enable_Adaptive_Cooling"
            });
            //Run control instance excepts the lastest loop. 
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.If,
                Parameter = "SrcCodeIndx<SrcCodeIndx1"
            });
            flowSheet.AddRows(GenerateIdsOffRows(idsRows));
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Test,
                Parameter = ControlAdaptiveCoolingInstanceName
            });
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.EndIf
            });
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Next
            });
            //When it ran with UFP, need to recover nWire.
            if (TestProgram.IgxlWorkBk.SubFlowSheets.Any(p => p.Value.Name.Equals($"{NwireSetting.ConFlownWire}Default_Enable", StringComparison.OrdinalIgnoreCase)))
            {
                IEnumerable<string> jobs = TestPlanStatic.JobInfoSheet?.Rows.Where(x => x.TesterType.ToUpper().Equals("UFP")).Select(x => x.JobName).ToList();
                flowSheet.AddRow(new FlowRow
                {
                    Job = jobs?.Any() == true ? string.Join(",", jobs) : "",
                    Opcode = jobs?.Any() == true ? OpCode.Call : OpCode.Nop,
                    Parameter = $"{NwireSetting.ConFlownWire}Default_Enable"
                });
            }
            else if (TestProgram.IgxlWorkBk.SubFlowSheets.Any(p => p.Value.Name.Equals($"{NwireSetting.ConFlownWire}Default", StringComparison.OrdinalIgnoreCase)))
            {
                IEnumerable<string> jobs = TestPlanStatic.JobInfoSheet?.Rows.Where(x => x.TesterType.ToUpper().Equals("UFP")).Select(x => x.JobName).ToList();
                flowSheet.AddRow(new FlowRow
                {
                    Job = jobs?.Any() == true ? string.Join(",", jobs) : "",
                    Opcode = jobs?.Any() == true ? OpCode.Call : OpCode.Nop,
                    Parameter = $"{NwireSetting.ConFlownWire}Default"
                });
            }
            //Relay off.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = noRelay ? OpCode.Nop : OpCode.Test,
                Parameter = "Relay_Off_HARDIP"
            });
            //Run clean instance in the end.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Test,
                Parameter = CleanUpAdaptiveCoolingInstanceName
            });
            //Bin out when adaptive cooling is enabled.
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.BinTable,
                Enable = "Enable_Adaptive_Cooling",
                Parameter = "Bin_TMPS_Monitor"
            });
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Print,
                Parameter = '"' + flowSheetName.Replace(" ", "_") + " Stop " + '"'
            });
            flowSheet.AddRow(new FlowRow
            {
                Opcode = OpCode.Return
            });

            return flowSheet;
        }

        protected List<FlowRow> GenerateTmpsBodyRows(List<FlowRow> rows, string flag = "", bool isBinOut = false)
        {
            var flowRows = new List<FlowRow>();
            foreach (FlowRow row in rows)
            {
                FlowRow copy = row.Copy();
                if (row.Opcode == OpCode.Test)
                {
                    copy.ColumnA = isBinOut ? "Relax Spex" : "Normal Spec";
                    copy.Enable = "";
                    copy.Job = "";
                }
                if (!string.IsNullOrEmpty(flag))
                {
                    copy.FailAction = flag;
                }
                else
                {
                    copy.FailAction = "";
                }
                if (row.GetJobs().Exists(x => x.Equals("cp2", StringComparison.CurrentCultureIgnoreCase) || x.Equals("ft2", StringComparison.CurrentCultureIgnoreCase)) && !string.IsNullOrEmpty(copy.HiLim) && copy.Units.Equals("C", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (LocalSpecs.Options.CoolingLimitValue.Split(',').Length > 2)
                    {
                        string limitVal = LocalSpecs.Options.CoolingLimitValue.Split(',')[1].Trim();
                        string binoutVal = LocalSpecs.Options.CoolingLimitValue.Split(',')[2].Trim();
                        copy.HiLim = isBinOut ? binoutVal : limitVal;
                    }
                }
                flowRows.Add(copy);
            }
            return flowRows;
        }

        private List<FlowRow> GenerateIdsOffRows(List<FlowRow> idsRows)
        {
            var flowRows = new List<FlowRow>();
            foreach (FlowRow row in idsRows)
            {
                FlowRow copy = row.Copy();
                List<string> tmpsIdsName = copy.Parameter.Split('_').ToList();
                tmpsIdsName.Insert(2, "TMPSMON");
                copy.Parameter = string.Join("_", tmpsIdsName);
                copy.FailAction = "";
                if (row.Opcode == OpCode.Test)
                {
                    copy.Enable = "";
                    copy.Job = "";
                }
                flowRows.Add(copy);
            }
            return flowRows;
        }

        private void GenerateTmpsInstanceIntoCommon()
        {
            var instanceRows = new List<InstanceRow>();

            Function initialFunction = TestProgram.VbtFunctionLib.GetFunctionByName(InitialAdaptiveCoolingInstanceName, "hardip", true);
            initialFunction.SetParamValue("indexName", $"{_indexName}");
            initialFunction.SetParamValue("startIndex", $"{_startIndex}");
            initialFunction.SetParamValue("endIndex", $"{_endIndex}");
            initialFunction.SetParamValue("flagName", $"{_flagName}");
            instanceRows.Add(new InstanceRow
            {
                TestName = InitialAdaptiveCoolingInstanceName,
                VbtType = initialFunction.Type,
                VbtName = initialFunction.FullFunctionName,
                ArgList = initialFunction.Parameters,
                Args = initialFunction.ArgList,
            });

            Function controlFunction = TestProgram.VbtFunctionLib.GetFunctionByName(ControlAdaptiveCoolingInstanceName, "hardip", true);
            instanceRows.Add(new InstanceRow
            {
                TestName = ControlAdaptiveCoolingInstanceName,
                VbtType = controlFunction.Type,
                VbtName = controlFunction.FullFunctionName,
                ArgList = controlFunction.Parameters,
                Args = controlFunction.ArgList,
            });

            Function cleanUpFunction = TestProgram.VbtFunctionLib.GetFunctionByName(CleanUpAdaptiveCoolingInstanceName, "hardip", true);
            instanceRows.Add(new InstanceRow
            {
                TestName = CleanUpAdaptiveCoolingInstanceName,
                VbtType = cleanUpFunction.Type,
                VbtName = cleanUpFunction.FullFunctionName,
                ArgList = cleanUpFunction.Parameters,
                Args = cleanUpFunction.ArgList,
            });

            InstanceSheet commonInstanceSheet;
            bool exist = false;
            string key = "";
            foreach (KeyValuePair<string, InstanceSheet> insSheet in TestProgram.IgxlWorkBk.InsSheets)
            {
                if (string.Equals(insSheet.Value.Name, OutputConst.RelayInstName, StringComparison.OrdinalIgnoreCase))
                {
                    exist = true;
                    key = insSheet.Key;
                    break;
                }
            }

            if (!exist)
            {
                commonInstanceSheet = new InstanceSheet(OutputConst.RelayInstName);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, commonInstanceSheet);
            }
            else
            {
                commonInstanceSheet = TestProgram.IgxlWorkBk.InsSheets[key];
            }
            commonInstanceSheet.AddRows(instanceRows);
        }
    }
}
