using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.Scan.Harvest.Flow;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow
{
    public class RtosFlowRowGenerator : HardIpFlowRowGenerator
    {
        public RtosFlowRowGenerator(HardIpInputData hardIpInputData, string sheetName) : base(hardIpInputData, sheetName)
        {
        }

        internal bool IdsNoFuse
        {
            get
            {
                return HardIpService.IdsNoFuse(Pat.MiscInfo);
            }
        }

        internal bool IsNandPattern
        {
            get
            {
                return HardIpService.IsNandPattern(Pat.Pattern.GetLastPayload());
            }
        }


        internal bool IsSpiPattern
        {
            get
            {
                return HardIpService.IsSpiPattern(Pat.Pattern.GetLastPayload());
            }
        }

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
        }

        public override List<FlowRow> GenRunScenarioRows(string rtosBootInstance = "RtosBoot", bool isFirstScenario = false)
        {
            var flowRows = new List<FlowRow>();
            string rtosTestName = CommonGenerator.GenRtosInsTestName(Pattern.FunctionName, BlockName, Pattern.SubBlock, Pat.Pattern.TestPlanPatternName) + "_" + LabelVoltage;
            string flagName = string.IsNullOrEmpty(Pat.Failflag) ? GetFlagName() : Pat.Failflag;
            Pat.Failflag = flagName;
            if (!isFirstScenario)
            {
                flowRows.Add(new FlowRow
                {
                    Enable = GenEnable("Rtos_" + LabelVoltage, ""),
                    Opcode = OpCode.Test,
                    Parameter = rtosBootInstance,
                    DeviceCondition = OpCode.FlagTrue,
                    DeviceName = flagName
                });
            }

            flowRows.Add(new FlowRow
            {
                Enable = GenEnable("Rtos_" + LabelVoltage, ""),
                Opcode = OpCode.FlagClear,
                Parameter = flagName,
                DeviceCondition = OpCode.FlagTrue,
                DeviceName = flagName
            });
            flowRows.Add(new FlowRow
            {
                Enable = GenEnable("Rtos_" + LabelVoltage, ""),
                Opcode = OpCode.Test,
                Parameter = rtosTestName,
                FailAction = flagName
            });

            return flowRows;
        }


        private string GetFlagName()
        {
            string flagName = "F_Rtos_func_";
            List<string> stringArray = Pat.Pattern.TestPlanPatternName.Split('_').ToList();
            string scenario = stringArray.Find(p => Regex.IsMatch(p, "SC", RegexOptions.IgnorePatternWhitespace));
            if (scenario == null)
            {
                flagName += Combination.CombineByUnderLine("SC", "SCCMD" + GetScenario());
            }
            else if (scenario.Equals("SC" + GetScenario()))
            {
                flagName += scenario;
            }
            else
            {
                flagName += Combination.CombineByUnderLine(scenario, "SCCMD" + GetScenario());
            }

            return flagName;
        }

        public string GetScenario()
        {
            string cmd = GetParaValue("cmd4");
            var regexRtos = new Regex(@"(?<Scenario>\d+\s+)", RegexOptions.IgnoreCase);
            Match match = regexRtos.Match(cmd);
            return match.Groups["Scenario"].ToString().Trim();
        }

        public string GetParaValue(string paraName)
        {
            string value = "";
            if (Pat.MiscInfoDict.ContainsKey(paraName.ToUpper()))
            {
                value = Pat.MiscInfoDict[paraName.ToUpper()];
            }

            return value;
        }

        public override List<FlowRow> GenTestRows(bool isCz2Only = false)
        {
            var flowRows = new List<FlowRow>();
            var hardipMainNonBincutInstance = new HardipNonBinCutInstance();
            var testFlowRow = new FlowRow
            {
                Opcode = CreateTestOpcode(),
                Parameter = CreateTestParameter(),
                Enable = GenEnable("Rtos_" + LabelVoltage, ""),
                FailAction = CreateTestFailAction()
            };
            //use limit rows
            List<FlowRow> limitRows = null;
            if (IdsNoFuse || IsNandPattern || IsSpiPattern)
            {
                limitRows = GenUseLimits(
                    testFlowRow.Parameter,
                    Pat,
                    LabelVoltage,
                    CreateUseLimitFailAction,
                    LocalSpecs.AllJobsRtos
                );
            }

            List<FlowRow> ifFlowRows = hardipMainNonBincutInstance.GetIfFlowRows(null, [testFlowRow], limitRows);
            flowRows.AddRange(ifFlowRows);


            return flowRows;
        }
        public override FlowRow GenBinTableRow(string voltage = "")
        {
            var flowRow = new FlowRow
            {
                Opcode = CreateBinTableOpcode(),
                Parameter = CreateBinTableParameter()
            };
            return flowRow;
        }

        protected override string CreateBinTableParameter()
        {
            if (!string.IsNullOrEmpty(Pat.FunctionName))
            {
                return string.IsNullOrEmpty(Pat.Failflag) ? $"{HardIpConstData.BinFlowFlag}_{BlockName}_{SubBlockName}" : Pat.Failflag.Replace("F_", "Bin_");
            }

            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""));
        }

        protected internal override string CreateTestParameter(bool forShmoo = false)
        {
            if (Pat.FunctionName.Equals(VbtFunctionLibShared.RtosBootUp, StringComparison.OrdinalIgnoreCase))
            {
                return BlockName + "_" + SubBlockName;
            }

            string patternName = Pat.Pattern.GetPatternName();

            if (!string.IsNullOrEmpty(Pat.TestName))
            {
                return Pat.TestName + "_" + ActualLabelVoltage;
            }

            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, patternName,
                Pat.PatternIndexFlag,
                TimingAc, Pat.ForceVoltageFlag, InstNameSubStr, "", NoPattern, Pat.WirelessData.IsNeedPostBurn, true, Pat.WirelessData.IsDoMeasure);
        }

        protected internal override string CreateTestFailAction()
        {
            if (Pat.MiscInfoDict.ContainsKey(HardIpConstData.IgnorePatBinOut) || ((HardIpConstData.RegOpcodeInPatt.IsMatch(Pat.Pattern.RealPatternName) && string.IsNullOrEmpty(Pat.Failflag))))
            {
                return "";
            }

            if (!string.IsNullOrEmpty(Pat.FunctionName))
            {
                return string.IsNullOrEmpty(Pat.Failflag) ? $"{HardIpConstData.PrefixHardIpFailAction}_{BlockName}_{SubBlockName}{HardIpConstData.SuffixHardIpFailAction}" : Pat.Failflag;
            }

            return CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), Pat.Pattern.GetLastPayload(), TimingAc,
                InstNameSubStr, "", Pat.MiscInfo, NoPattern);
        }
        internal string CreateUseLimitFailAction(string repeatStr)
        {
            if (Pat.MiscInfoDict.ContainsKey(HardIpConstData.IgnorePatBinOut))
            {
                return "";
            }

            if (IdsNoFuse)
            {
                return string.IsNullOrEmpty(Pat.Failflag) ? $"{HardIpConstData.PrefixHardIpFailAction}_{BlockName}_{SubBlockName}{HardIpConstData.SuffixHardIpFailAction}" : Pat.Failflag;
            }

            if (repeatStr != "")
            {
                return "F_IDS_Current_Main" + "_" + repeatStr.Replace(".", "p");
            }

            return "F_IDS_Current_Main_1x";
        }
    }
}
