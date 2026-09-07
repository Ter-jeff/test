using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;
using Automation.Utility.HardIP;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness
{
    public class InitFlagGenerator
    {
        private static readonly string _initFlowSheetName = "Flow_HardIP_Init_Flag";
        private readonly FlowRow _printStartRow = new FlowRow { Opcode = "print", Parameter = '"' + "Flow " + _initFlowSheetName.Replace("Flow_", "") + " Start" + '"' };
        private readonly FlowRow _printStopRow = new FlowRow { Opcode = "print", Parameter = '"' + "Flow " + _initFlowSheetName.Replace("Flow_", "") + " Stop" + '"' };
        private readonly FlowRow _returnRow = new FlowRow { Opcode = "return" };
        public SubFlowSheet GenInitFlag(Dictionary<string, HardIpSheet> planDic, BinTableSheet binTableSheet)
        {
            string dir = LocalSpecs.IsBenchLog ? FolderStructure.TarDir : FolderStructure.DirHardIp;
            var initFlow = new SubFlowSheet(_initFlowSheetName, "Initialize");
            SubFlowSheet existInitFlow = TestProgram.IgxlWorkBk.GetFlowSheet(_initFlowSheetName, dir);
            if (existInitFlow.Rows.Any())
            {
                initFlow = RemoveReturnFlows(existInitFlow);
            }
            else
            {
                #region Print Start
                initFlow.AddRow(_printStartRow);
                #endregion
            }
            #region generate bintable flag list for checking

            List<string> bintableFList = new List<string>();
            foreach (BinTableRow row in binTableSheet.Rows)
            {
                bintableFList.AddRange(row.ItemList.Split(',').ToList());
            }
            #endregion

            bintableFList = bintableFList.Distinct().ToList();
            List<string> initflowFList = new List<string>();
            foreach (KeyValuePair<string, HardIpSheet> sheet in planDic)
            {
                if (Regex.IsMatch(sheet.Key, "DCTEST_IDS", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                var patternNames = new List<string>();
                foreach (HardIpPattern pattern in sheet.Value.Rows)
                {
                    string patternName = pattern.Pattern.GetPatternName();
                    string blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
                    string subBlockName = CommonGenerator.GetSubBlockName(patternName, pattern.MiscInfo, blockName);

                    #region Non-IDS sheet

                    string paramter;
                    string timingAc = CommonGenerator.GetTimingAc(pattern);
                    string instNameSubStr = HardIpService.GetInstNameSubStr(pattern.MiscInfo);
                    bool noPattern = HardIpService.IsNoPattern(pattern.Pattern.GetLastPayload());

                    //add subBlockName for checking repeated fail flag by Raze, solve different subblock but not gen the fail flag issue 2017/6/12

                    if (timingAc != "")
                    {
                        paramter = "Bin_" + blockName + "_" + subBlockName + "_" + patternName + "_" + timingAc;
                        //failFlag = "F_" + blockName + "_" + pattern.PatternName + "_" + timingAc + "_Flag";
                    }
                    else
                    {
                        paramter = "Bin_" + blockName + "_" + subBlockName + "_" + patternName;
                        //failFlag = "F_" + blockName + "_" + pattern.PatternName + "_Flag";
                    }
                    string failFlagN = CommonGenerator.GenHardIpFlowFailFlag(pattern.SheetName, blockName, subBlockName.ToUpper().Replace("-MERGE", ""), patternName.Replace(PatternClass.MultipleConst, ""), timingAc, instNameSubStr, HardIpConstData.LabelNv, noPattern);
                    string failFlagH = CommonGenerator.GenHardIpFlowFailFlag(pattern.SheetName, blockName, subBlockName.ToUpper().Replace("-MERGE", ""), patternName.Replace(PatternClass.MultipleConst, ""), timingAc, instNameSubStr, HardIpConstData.LabelHv, noPattern);
                    string failFlagL = CommonGenerator.GenHardIpFlowFailFlag(pattern.SheetName, blockName, subBlockName.ToUpper().Replace("-MERGE", ""), patternName.Replace(PatternClass.MultipleConst, ""), timingAc, instNameSubStr, HardIpConstData.LabelLv, noPattern);
                    if (patternNames.Contains(paramter))
                    {
                        continue;
                    }

                    patternNames.Add(paramter);
                    initflowFList.Add(failFlagN);
                    initflowFList.Add(failFlagH);
                    initflowFList.Add(failFlagL);
                    #endregion
                }
            }
            initflowFList = initflowFList.Distinct().ToList();
            #region check fail-flag referred to bintable if generated init-flag flow not exist
            var initflowFListTmp = new List<string>();
            foreach (string binflag in bintableFList)
            {
                if (!initflowFList.Contains(binflag))
                {
                    initflowFListTmp.Add(binflag);
                }
            }
            #endregion
            initflowFList.AddRange(initflowFListTmp);

            foreach (IGrouping<string, string> flags in initflowFList.GroupBy(x => x.Split('_')[1] + "&&" + x.Split('_')[2]).ToList())
            {
                var flagTmp = new List<string>();
                foreach (string flag in flags)
                {
                    flagTmp.Add(flag);
                    if (string.Join(",", flagTmp).Length >= 6000)
                    {
                        initFlow.AddRow(new FlowRow { Opcode = OpCode.FlagClear, Parameter = string.Join(",", flagTmp) });
                        flagTmp = new List<string>();
                    }
                }
                if (flagTmp.Any())
                {
                    initFlow.AddRow(new FlowRow { Opcode = OpCode.FlagClear, Parameter = string.Join(",", flagTmp) });
                }
            }

            #region Print Stop
            initFlow.AddRow(_printStopRow);
            #endregion

            #region Print return
            initFlow.AddRow(_returnRow);
            #endregion
            return initFlow;
        }
        private SubFlowSheet RemoveReturnFlows(SubFlowSheet existInitFlow)
        {
            List<FlowRow> removeRowList = new List<FlowRow> { _printStopRow, _returnRow };
            foreach (FlowRow removeRow in removeRowList)
            {
                int removeIndex = existInitFlow.Rows.FindIndex(x => x.Opcode.Equals(removeRow.Opcode) && x.Parameter.Equals(removeRow.Parameter));
                if (removeIndex != -1)
                {
                    existInitFlow.Rows.RemoveAt(removeIndex);
                }
            }
            return existInitFlow;
        }
    }
}
