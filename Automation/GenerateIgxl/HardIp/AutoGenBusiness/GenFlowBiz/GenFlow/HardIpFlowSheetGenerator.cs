using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.DividerManager;
using Automation.GenerateIgxl.HardIp.DividerManager.FlowDividerManager;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using LogLib.Utility;

using ScghLib.Reader;

using TestPlanLib.BinNumberLegacy;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow
{
    public class HardIpFlowSheetGenerator : FlowSheetGeneratorBase
    {
        protected string Volatage = "";

        public HardIpFlowSheetGenerator(HardIpInputData hardIpInputData, string sheetName, List<HardIpPattern> patternList = null) : base(hardIpInputData, sheetName, patternList)
        {
            FlowRowGenerator = new HardIpFlowRowGenerator(hardIpInputData, sheetName);
        }

        protected override List<HardIpPattern> DividePatterns()
        {
            var dividedPatList = new List<HardIpPattern>();
            foreach (HardIpPattern pattern in PatternList)
            {
                try
                {
                    bool isHardIpUniversal = SearchInfo.GetVbtNameByPattern(HardIpInputData, pattern) == "";
                    if ((HardIpConstData.RegInsInPatt.IsMatch(pattern.Pattern.RealPatternName) ||
                         HardIpConstData.RegOpcodeInPatt.IsMatch(pattern.Pattern.RealPatternName)) &&
                        pattern.MeasPins.Count == 0)
                    {
                        dividedPatList.Add(pattern);
                        continue;
                    }

                    List<HardIpPattern> tempList1 = DividerMain.DivideMeasPins(HardIpInputData, pattern, isHardIpUniversal, true);
                    List<HardIpPattern> tempList4 = new FlowLimitDivider().DivideUseLimit(tempList1, HardIpInputData);
                    dividedPatList.AddRange(tempList4);
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                    throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum);
                }
            }
            return dividedPatList;
        }

        internal string GetFlowSequence(string voatlge)
        {
            if (voatlge.Equals("NV", StringComparison.CurrentCultureIgnoreCase))
            {
                return HardIpConstData.LabelNv;
            }

            if (voatlge.Equals("LV", StringComparison.CurrentCultureIgnoreCase))
            {
                return HardIpConstData.LabelLv;
            }

            if (voatlge.Equals("HV", StringComparison.CurrentCultureIgnoreCase))
            {
                return HardIpConstData.LabelHv;
            }

            return "";
        }

        protected override List<FlowRow> GenFlowBodyRows(bool shmooflag = false, bool vtShmooFlag = false)
        {
            var flowBodyRows = new List<FlowRow>();

            List<string> volatages = new List<string>();
            if (Volatage != "")
            {
                volatages.Add(Volatage);
            }
            else
            {
                volatages.Add(GetFlowSequence(HardIpConstData.LabelNv));
                if (!SheetName.Equals("INIT", StringComparison.CurrentCultureIgnoreCase))
                {
                    volatages.Add(GetFlowSequence(HardIpConstData.LabelLv));
                    volatages.Add(GetFlowSequence(HardIpConstData.LabelHv));
                }
            }

            foreach (string votlage in volatages)
            {
                if (vtShmooFlag)
                {
                    if (!votlage.ContainsIgnoreCase("NV"))
                    {
                        continue;
                    }

                    flowBodyRows.AddRange(GenFlowTestRowsByVoltage("NV", shmooflag, true));
                }
                else
                {
                    flowBodyRows.AddRange(GenFlowTestRowsByVoltage(votlage, shmooflag, false));
                }
            }


            if (!vtShmooFlag && !shmooflag)
            {
                if (!LocalSpecs.Options.BinTableBeforeEachItem)
                {
                    flowBodyRows.AddRange(GenFlowBinTableRows());
                }

                flowBodyRows.AddRange(FlowRowGenerator.GenTtrFlagClearRow(flowBodyRows));
            }

            return flowBodyRows;
        }

        protected virtual List<FlowRow> GenFlowTestRowsByVoltage(string labelVoltage, bool shmooCharflag, bool vtShmooflag)
        {
            var allflowRows = new List<FlowRow>();
            var flowRows = new List<FlowRow>();

            FlowRowGenerator.LabelVoltage = labelVoltage;
            FlowRowGenerator.Init3XFlags.Clear();
            FlowRowGenerator.Init3XEnableWds.Clear();

            string previousLoopGroupStatus = "";
            var loopGroupNextRows = new List<FlowRow>();

            foreach (HardIpPattern pattern in ExtendedPatList)
            {
                try
                {
                    GenFlowTestRowsByPattern(pattern, labelVoltage, shmooCharflag, vtShmooflag, flowRows, ref previousLoopGroupStatus, ref loopGroupNextRows);
                }
                catch (Exception ex)
                {
                    ErrorMessageBox.Show(string.Format(ex.ToString()));
                    throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum);
                }
            }
            if (loopGroupNextRows != null)
            {
                flowRows.AddRange(loopGroupNextRows);
                loopGroupNextRows.Clear();
            }

            if (flowRows.Count != 0)
            {
                flowRows.AddRange(GenResetRelayRows(labelVoltage));
            }

            foreach (KeyValuePair<string, string> initItem in FlowRowGenerator.InitOriginalItems)
            {
                allflowRows.Add(new FlowRow { Opcode = OpCode.FlagClear, Parameter = initItem.Value });
                allflowRows.Add(new FlowRow { Enable = initItem.Key, Opcode = OpCode.Nop });
                allflowRows.Add(new FlowRow { Opcode = OpCode.DisableFlowWd, Parameter = initItem.Key });
            }
            FlowRowGenerator.InitOriginalItems.Clear();
            allflowRows.AddRange(flowRows);

            return allflowRows;
        }

        private void GenFlowTestRowsByPattern(HardIpPattern pattern, string labelVoltage, bool shmooCharflag, bool vtShmooflag,
            List<FlowRow> flowRows, ref string previousLoopGroupStatus, ref List<FlowRow> loopGroupNextRows)
        {
            string loopGroupStatus = pattern.MiscInfo.Split(';').FirstOrDefault(x => x.StartsWith("LoopGroup", StringComparison.OrdinalIgnoreCase)) ??
                                     "";

            if (pattern.SkipList.Contains(labelVoltage) || pattern.SkipDotNet)
            {
                return;
            }

            if (!IsNeedGenerate(pattern))
            {
                return;
            }

            FlowRowGenerator.Pat = pattern;
            if (FlowRowGenerator.NoNeedToGen)
            {
                return;
            }

            List<FlowRow> sweepShmooRows = FlowRowGenerator.GenShmooRows(labelVoltage);
            if (shmooCharflag && !pattern.ForceCondition.IsShmooInCharFlow)
            {
                return;
            }

            if (vtShmooflag && !pattern.ForceCondition.IsVtShmoo)
            {
                return;
            }

            List<FlowRow> sweepCodeForRow = FlowRowGenerator.GenSweepCodeForRow();
            List<FlowRow> sweepVoltageRows = FlowRowGenerator.GenSweepVoltageForRow();
            FlowRow retestIfRow = FlowRowGenerator.GenRetestIfRow();
            FlowRow retestEndIfRow = FlowRowGenerator.GenRetestEndIfRow();
            List<FlowRow> sweepCodeNextRow = FlowRowGenerator.GenSweepCodeOrVoltageNextRow();
            bool isSingleInstanceUseSweep = false;
            if (!string.IsNullOrEmpty(previousLoopGroupStatus) &&
                !loopGroupStatus.Equals(previousLoopGroupStatus, StringComparison.OrdinalIgnoreCase)) //Generate "Next" of loop group at the top of next item.
            {
                if (loopGroupNextRows != null)
                {
                    flowRows.AddRange(loopGroupNextRows);
                    loopGroupNextRows.Clear();
                }
            }

            if (shmooCharflag) //for Flow_HardIP_Char
            {
                GenCharFlowRowsByPattern(pattern, labelVoltage, loopGroupStatus, previousLoopGroupStatus, sweepShmooRows, sweepCodeForRow, sweepVoltageRows, sweepCodeNextRow, retestIfRow, flowRows, ref isSingleInstanceUseSweep, ref loopGroupNextRows);
            }
            else               //For Flow_HardIP_XXXX   
            {
                GenProdFlowRowsByPattern(pattern, labelVoltage, shmooCharflag, vtShmooflag, loopGroupStatus, previousLoopGroupStatus, sweepShmooRows, sweepCodeForRow, sweepVoltageRows, sweepCodeNextRow, retestIfRow, flowRows, ref isSingleInstanceUseSweep, ref loopGroupNextRows);
            }

            if (retestEndIfRow != null)
            {
                flowRows.Add(retestEndIfRow);
            }

            if (string.IsNullOrEmpty(loopGroupStatus)) //Item without loop group generate "Next"
            {
                if (sweepCodeNextRow != null && isSingleInstanceUseSweep)
                {
                    flowRows.AddRange(sweepCodeNextRow);
                }
            }
            flowRows.AddRange(FlowRowGenerator.GenOpcodeRowsAftPat());

            //Record the previous loop gorup status
            previousLoopGroupStatus = loopGroupStatus;
            if (FlowRowGenerator.NeedRealtimeBinOut)
            {
                flowRows.AddRange(GenFlowSinPatBinTableRows(pattern));
            }
        }

        private void GenCharFlowRowsByPattern(HardIpPattern pattern, string labelVoltage, string loopGroupStatus, string previousLoopGroupStatus,
            List<FlowRow> sweepShmooRows, List<FlowRow> sweepCodeForRow, List<FlowRow> sweepVoltageRows, List<FlowRow> sweepCodeNextRow,
            FlowRow retestIfRow, List<FlowRow> flowRows, ref bool isSingleInstanceUseSweep, ref List<FlowRow> loopGroupNextRows)
        {
            if (pattern.ForceCondition.IsShmooInCharFlow)
            {

                flowRows.AddRange(FlowRowGenerator.GenExtraRowsByMisc());
                if (IsNeedGenerateRelay(pattern))
                {
                    flowRows.AddRange(FlowRowGenerator.GenRelayRows());
                }

                flowRows.AddRange(FlowRowGenerator.GenNwireChangeRows());
                flowRows.AddRange(FlowRowGenerator.GenNwireDisOrEnableRows());

                List<string> opcodeBeforeList = SearchInfo.GetOpcode(pattern, "B");
                CommonGenerator.ConvertPatNameInOpcode(opcodeBeforeList, flowRows, labelVoltage);
                flowRows.AddRange(FlowRowGenerator.GenOpcodeRowsBefPat(opcodeBeforeList));
                if (!(!string.IsNullOrEmpty(loopGroupStatus) &&
                    loopGroupStatus.Equals(previousLoopGroupStatus, StringComparison.OrdinalIgnoreCase))) //The middle items of loop group don't generate "For"
                {
                    if (sweepCodeForRow != null)
                    {
                        flowRows.AddRange(sweepCodeForRow);
                        isSingleInstanceUseSweep = true;
                    }
                    if (sweepVoltageRows != null)
                    {
                        flowRows.AddRange(sweepVoltageRows);
                        isSingleInstanceUseSweep = true;
                    }
                    if (!string.IsNullOrEmpty(loopGroupStatus))
                    {
                        loopGroupNextRows = sweepCodeNextRow;
                    }
                }
                flowRows.AddRange(FlowRowGenerator.GenPreRetestRows());
                if (retestIfRow != null)
                {
                    flowRows.Add(retestIfRow);
                }

                flowRows.AddRange(FlowRowGenerator.GenTestRows(true));
                if (pattern.ForceCondition.IsShmooInForce)
                {
                    flowRows.AddRange(sweepShmooRows);
                }
            }
        }

        private void GenProdFlowRowsByPattern(HardIpPattern pattern, string labelVoltage, bool shmooCharflag, bool vtShmooflag,
            string loopGroupStatus, string previousLoopGroupStatus, List<FlowRow> sweepShmooRows, List<FlowRow> sweepCodeForRow,
            List<FlowRow> sweepVoltageRows, List<FlowRow> sweepCodeNextRow, FlowRow retestIfRow, List<FlowRow> flowRows,
            ref bool isSingleInstanceUseSweep, ref List<FlowRow> loopGroupNextRows)
        {
            if (vtShmooflag && pattern.ForceCondition.IsVtShmoo)
            {
                flowRows.AddRange(FlowRowGenerator.GenExtraRowsByMisc());
                if (IsNeedGenerateRelay(pattern))
                {
                    flowRows.AddRange(FlowRowGenerator.GenRelayRows());
                }

                flowRows.AddRange(FlowRowGenerator.GenNwireChangeRows());
                flowRows.AddRange(FlowRowGenerator.GenNwireDisOrEnableRows());

                flowRows.AddRange(sweepShmooRows);
            }
            else if (pattern.ForceCondition.IsShmooInProdFlow)
            {
                flowRows.AddRange(FlowRowGenerator.GenExtraRowsByMisc());
                if (IsNeedGenerateRelay(pattern))
                {
                    flowRows.AddRange(FlowRowGenerator.GenRelayRows());
                }

                flowRows.AddRange(FlowRowGenerator.GenNwireChangeRows());
                flowRows.AddRange(FlowRowGenerator.GenNwireDisOrEnableRows());

                List<string> opcodeBeforeList = SearchInfo.GetOpcode(pattern, "B");
                CommonGenerator.ConvertPatNameInOpcode(opcodeBeforeList, flowRows, labelVoltage);
                if (!shmooCharflag && pattern.ForceCondition.IsShmooInProdFlow)
                {
                    flowRows.AddRange(FlowRowGenerator.GenOpcodeRowsBefPat(opcodeBeforeList));
                }

                if (!(!string.IsNullOrEmpty(loopGroupStatus) && loopGroupStatus.Equals(previousLoopGroupStatus, StringComparison.OrdinalIgnoreCase)))
                //Generate "For" except the middle items of loop group
                {
                    if (sweepCodeForRow != null)
                    {
                        flowRows.AddRange(sweepCodeForRow);
                        isSingleInstanceUseSweep = true;
                    }
                    if (sweepVoltageRows != null)
                    {
                        flowRows.AddRange(sweepVoltageRows);
                        isSingleInstanceUseSweep = true;
                    }
                    if (!string.IsNullOrEmpty(loopGroupStatus))
                    {
                        loopGroupNextRows = sweepCodeNextRow;
                    }
                }
                flowRows.AddRange(FlowRowGenerator.GenPreRetestRows());
                if (retestIfRow != null)
                {
                    flowRows.Add(retestIfRow);
                }

                flowRows.AddRange(FlowRowGenerator.GenTestRows());
            }
        }

        private List<FlowRow> GenFlowBinTableRows()
        {
            if (LocalSpecs.IsPatternValidate)
            {
                return new List<FlowRow>();
            }
            var flowBinTableRows = new List<FlowRow>();
            var duplicateFlow = new List<string>();
            foreach (HardIpPattern pattern in ExtendedPatList)
            {
                FlowRowGenerator.Pat = pattern;
                if (VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)) ||
                    pattern.IsIgnorePatBinOut() ||
                    Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^cz_", RegexOptions.IgnoreCase) ||
                    (HardIpConstData.RegOpcodeInPatt.IsMatch(pattern.Pattern.RealPatternName) && string.IsNullOrEmpty(pattern.Failflag)) ||
                    FlowRowGenerator.Pat.MiscInfoDict.ContainsKey(HardIpConstData.RealtimePatBinOut))
                {
                    continue;
                }

                FlowRow binRow = FlowRowGenerator.GenBinTableRow();
                var binRows = new List<FlowRow> { binRow };

                if (!HardIpConstData.RegInsInPatt.IsMatch(pattern.Pattern.RealPatternName) &&
                    !HardIpConstData.RegOpcodeInPatt.IsMatch(pattern.Pattern.RealPatternName))
                {
                    ProdCharSheetRow scghRow = ScghStatic.ScghData.GetProdCharSheetRow(pattern.Pattern.GetLastPayload());
                    binRows = SplitBinRowByVoltage(binRow, SearchInfo.GetHardIpBinRangeItem(pattern, scghRow));
                }

                foreach (FlowRow row in binRows)
                {
                    if (!duplicateFlow.Contains(row.Parameter))
                    {
                        flowBinTableRows.Add(row);
                        duplicateFlow.Add(row.Parameter);
                    }
                }
            }

            return flowBinTableRows;
        }

        private List<FlowRow> GenFlowSinPatBinTableRows(HardIpPattern pattern)
        {
            if (LocalSpecs.IsPatternValidate ||
                VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)) ||
                pattern.IsIgnorePatBinOut() ||
                Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^cz_", RegexOptions.IgnoreCase))
            {
                return new List<FlowRow>();
            }

            var flowBinTableRows = new List<FlowRow>();
            FlowRowGenerator.Pat = pattern;

            FlowRow binRow = FlowRowGenerator.GenBinTableRow();
            var binRows = new List<FlowRow> { binRow };

            if (!HardIpConstData.RegInsInPatt.IsMatch(pattern.Pattern.GetLastPayload()))
            {
                ProdCharSheetRow scghRow = ScghStatic.ScghData.GetProdCharSheetRow(pattern.Pattern.GetLastPayload());
                binRows = SplitBinRowByVoltage(binRow, SearchInfo.GetHardIpBinRangeItem(pattern, scghRow));
            }

            foreach (FlowRow row in binRows)
            {
                flowBinTableRows.Add(row);
            }

            return flowBinTableRows;
        }

        internal List<FlowRow> SplitBinRowByVoltage(FlowRow row, SoftBinRangeData volLabels)
        {
            List<string> labelVol = new List<string> { "H", "L", "N" };
            var binRows = new List<FlowRow>();
            if (volLabels.HardHlvBin != "")
            {
                FlowRow binHl = row.Copy();
                binHl.Parameter += "_HLV";
                binRows.Add(binHl);
            }
            if (volLabels.HardHvBin != "")
            {
                FlowRow binH = row.Copy();
                binH.Parameter += "_HV";
                binRows.Add(binH);
                labelVol.Remove("H");
            }
            if (volLabels.HardLvBin != "")
            {
                FlowRow binL = row.Copy();
                binL.Parameter += "_LV";
                binRows.Add(binL);
                labelVol.Remove("L");
            }
            if (volLabels.HardNvBin != "")
            {
                FlowRow binN = row.Copy();
                binN.Parameter += "_NV";
                binRows.Add(binN);
                labelVol.Remove("N");
            }
            if (labelVol.Count > 0)
            {
                binRows.Add(row);
            }

            return binRows;
        }

        private bool IsNeedGenerateRelay(HardIpPattern pattern)
        {
            return !(pattern == ExtendedPatList.First() && pattern.NewRelaySetting.Count == 0);
        }

        internal bool IsNeedGenerate(HardIpPattern pattern)
        {
            if (!string.IsNullOrEmpty(pattern.BlockType))
            {
                if (pattern.MeasPins.Count == 0)
                {
                    if (pattern.Pattern.InstancePayloadName.Count == 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
    }
}
