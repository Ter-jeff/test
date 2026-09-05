using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.BistBira.Base;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using ScghLib.Enums;
using ScghLib.Reader;
using ScghLib.Utility;

using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.BistBira
{
    public class MbistFlowGenerator
    {
        private readonly BistProdFlowSheet _prodFlowSheet;
        private readonly MbistBinTableType _binTableType;
        private readonly BistNaming _bistNaming;
        private readonly string _conMbist = "Mbist";
        private readonly string _module;
        private readonly string _sheetName;
        private List<string> _primeFlagList = new List<string>();
        private const string MbistEfuse = "MbistEfuse";
        private SubFlowSheet _subFlowSheet;
        private readonly List<string> _loopSyntax = new List<string> { "firstLoopNum", "secondLoopNum" };
        private readonly List<string> _usedFlagList = new List<string>();
        private readonly Dictionary<string, List<string>> _domainPrimeFlag = new Dictionary<string, List<string>>();
        private readonly bool _isCof;
        private readonly bool _mbistLoop;
        private readonly Dictionary<string, List<BistProdFlowRow>> _labelDic;
        private readonly Dictionary<string, List<BistProdFlowRow>> _passBranchDic;
        private readonly Dictionary<string, List<BistProdFlowRow>> _failBranchDic;

        public MbistFlowGenerator(BistProdFlowSheet prodFlowSheet, MbistConfig pConfig, MbistBinTableType binTableType, bool mbistLoop)
        {
            _prodFlowSheet = prodFlowSheet;
            _sheetName = prodFlowSheet.MbistSheet.SheetName;
            _bistNaming = new BistNaming(pConfig);
            _module = _bistNaming.GetModule(_sheetName);
            _isCof = prodFlowSheet.MbistSheet.IsCof;
            _labelDic = prodFlowSheet.LabelDic;
            _binTableType = binTableType;
            _passBranchDic = prodFlowSheet.Rows.GroupBy(x => x.PassBranch).ToDictionary(x => x.Key, x => x.ToList(), StringComparer.OrdinalIgnoreCase);
            _failBranchDic = prodFlowSheet.Rows.GroupBy(x => x.FailBranch).ToDictionary(x => x.Key, x => x.ToList(), StringComparer.OrdinalIgnoreCase);
            _mbistLoop = mbistLoop;
        }

        public List<BinTableRow> CreateBinTable(List<BinTableRow> needBinoutList, bool isNeedEvsDeferredBinout)
        {
            var bintables = new List<BinTableRow>();
            for (int i = 0; i < needBinoutList.Count; i++)
            {
                BinTableRow bintableRow = needBinoutList[i];
                BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Mbist", _module, "", bintableRow);
                bintableRow.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                bintableRow.Sort = binNumInfo.SoftBin.ToString("G15");
                bintableRow.Result = "Fail";
                if (isNeedEvsDeferredBinout)
                {
                    bintables.Add(AddDeferBinTable(ref bintableRow));
                }

                bintables.Add(bintableRow);
            }
            return bintables;
        }

        private BinTableRow AddDeferBinTable(ref BinTableRow row)
        {
            BinTableRow deferBinTable = row.Copy();
            deferBinTable.ItemList += ",F_EVS_Defer";
            deferBinTable.Items.Add("T");
            BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("EVSDEFER", _module, "MBIST", deferBinTable);
            deferBinTable.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
            deferBinTable.Sort = binNumInfo.SoftBin.ToString("G15");
            row.ItemList = "";
            return deferBinTable;
        }

        public (SubFlowSheet, List<BinTableRow>) GenerateSubflow()
        {
            SheetNameInit();
            PassBranchCheck();
            var subFlow = new List<FlowRow>();
            List<BinTableRow> needBinoutList = new List<BinTableRow>();
            var priousRow = new BistProdFlowRow();
            int loopIndex = 0;
            var controlFlagInitRows = new List<FlowRow>();
            var flagInitRows = new List<FlowRow>();
            subFlow.Add(new FlowRow { Opcode = "print", Parameter = "\"Flow " + _sheetName + " Start \"" });
            var loopSyntaxStack = new Stack<string>();
            #region Flags initialize
            InitializeFlags(controlFlagInitRows, flagInitRows);
            subFlow.AddRange(controlFlagInitRows);
            subFlow.AddRange(flagInitRows);
            #endregion
            Dictionary<string, string> devConRf = new Dictionary<string, string>();
            foreach (BistProdFlowRow row in _prodFlowSheet.Rows)
            {
                var flowRows = new FlowRows();
                if (BistAction.GetActionType(row) == BistActionType.RunPattern)
                {
                    AddRunPatternRows(row, flowRows);
                }
                if (BistAction.GetActionType(row) == BistActionType.LoopStart)
                {
                    AddLoopStartRows(row, flowRows, controlFlagInitRows, loopSyntaxStack, ref loopIndex);
                }
                if (BistAction.GetActionType(row) == BistActionType.LoopEnd)
                {
                    AddLoopEndRows(row, flowRows, loopSyntaxStack);
                }
                if (BistAction.GetActionType(row) == BistActionType.Retention)
                {
                    flowRows.Add(AddAction(row, GenerateRetention(priousRow, row)));
                }
                if (BistAction.GetActionType(row) == BistActionType.Get)
                {
                    AddGetRows(row, flowRows);
                }
                if (BistAction.GetActionType(row) == BistActionType.Set)
                {
                    AddSetRows(row, flowRows);
                }
                if (BistAction.GetActionType(row) == BistActionType.SetDefault)
                {
                    AddSetDefaultRows(row, flowRows);
                }
                if (BistAction.GetActionType(row) == BistActionType.Fail)
                {
                    AddFailRows(row, flowRows, needBinoutList);
                }
                if (BistAction.GetActionType(row) == BistActionType.Null)
                {
                    if (row.PassBranch == row.FailBranch && LocalSpecs.Options.Device != EnumDevice.RF)
                    {
                        flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = GetBranchByLabel(row.PassBranch) }));
                    }
                }
                if (BistAction.GetActionType(row) == BistActionType.Pass)
                {
                    flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.Return });
                }
                priousRow = row;
                ApplyDeviceCondition(row, flowRows, devConRf);

                subFlow.AddRange(flowRows);
            }
            _subFlowSheet.AddRows(subFlow);
            AddRepairJudgementByDomain(_subFlowSheet);
            AddRepairJudgement(_subFlowSheet);
            var printEndRow = new FlowRow { Opcode = "print", Parameter = "\"Flow " + _sheetName + " End \"" };
            _subFlowSheet.InsertBeforeReturnRow(new List<FlowRow> { printEndRow });
            if (_isCof)
            {
                foreach (FlowRow row in _subFlowSheet.Rows)
                {
                    row.DeviceName = "";
                    row.DeviceCondition = "";
                }
            }
            return (_subFlowSheet, needBinoutList);
        }

        private void InitializeFlags(List<FlowRow> controlFlagInitRows, List<FlowRow> flagInitRows)
        {
            if (LocalSpecs.Options.Device != EnumDevice.RF)
            {
                bool isFirstLabel = true;
                string para = "";
                var usedLabelList = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var usedFlagList = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (BistProdFlowRow row in _prodFlowSheet.Rows)
                {
                    if (usedLabelList.Contains(row.Label.ToUpper()))
                    {
                        continue;
                    }

                    if (isFirstLabel)
                    {
                        para += "Control_" + row.Label;
                        isFirstLabel = false;
                    }
                    else
                    {
                        para += ",Control_" + row.Label;
                    }

                    usedLabelList.Add(row.Label.ToUpper());
                    if (para.Length >= 4000)
                    {
                        controlFlagInitRows.Add(new FlowRow { Opcode = OpCode.FlagClear, Parameter = para });
                        isFirstLabel = true;
                        para = "";
                    }
                }
                if (!string.IsNullOrEmpty(para))
                {
                    controlFlagInitRows.Add(new FlowRow { Opcode = OpCode.FlagClear, Parameter = para });
                }
                para = "";
                isFirstLabel = true;

                foreach (BistProdFlowRow row in _prodFlowSheet.Rows)
                {
                    CheckPatternType(row, out bool _, out bool _, out bool isRepair);
                    if (CheckBranchType(row.FailBranch) != BistActionType.Fail && !isRepair)
                    {
                        continue;
                    }

                    string flag = "F_" + GetFailFlagByBranch(row);

                    row.FailFlag = flag;
                    if (usedFlagList.Contains(flag))
                    {
                        continue;
                    }

                    if (isFirstLabel)
                    {
                        para += flag;
                        isFirstLabel = false;
                    }
                    else
                    {
                        para += "," + flag;
                    }

                    if (para.Length >= 4000)
                    {
                        flagInitRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = para });
                        isFirstLabel = true;
                        para = "";
                    }
                    usedFlagList.Add(flag);
                }
                if (!string.IsNullOrEmpty(para))
                {
                    flagInitRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = para });
                }
            }
        }

        private void AddRunPatternRows(BistProdFlowRow row, FlowRows flowRows)
        {
            var flowRow = new FlowRow();
            CheckPatternType(row, out bool isEfp, out bool isRbox, out bool isRepair);
            flowRow.Parameter = _bistNaming.CreateNewTestName(_module, row);
            flowRow.Opcode = OpCode.Test;
            flowRow.ColumnA = row.Label;
            flowRow = AddCondition(row, flowRow);
            flowRow = AddAction(row, flowRow, isRepair);
            if (isEfp || isRbox)
            {
                flowRow.Enable = MbistEfuse;
            }

            flowRows.Add(flowRow);

            if (!_mbistLoop)
            {
                AddSkipNextControlRow(row, flowRows);
            }

            if (isEfp)
            {
                flowRows.Add(new FlowRow { Enable = "!" + MbistEfuse, ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = "Control_" + row.PassBranch, DeviceCondition = OpCode.FlagTrue, DeviceName = "Control_" + row.Label });
            }
        }

        private void AddSkipNextControlRow(BistProdFlowRow row, FlowRows flowRows)
        {
            if (!(row.PassBranch.ToUpper() == "NEXT" && row.FailBranch.ToUpper() == "NEXT"))
            {
                if (row.PassBranch.ToUpper() == "NEXT")
                {
                    flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagFalse, Parameter = "Control_" + row.Label, DeviceName = GetBranchByLabel(row.FailBranch), DeviceCondition = OpCode.FlagTrue });
                }
                else if (row.FailBranch.ToUpper() == "NEXT")
                {
                    flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagFalse, Parameter = "Control_" + row.Label, DeviceName = GetBranchByLabel(row.PassBranch), DeviceCondition = OpCode.FlagTrue });
                }
            }
        }

        private void AddLoopStartRows(BistProdFlowRow row, FlowRows flowRows, List<FlowRow> controlFlagInitRows, Stack<string> loopSyntaxStack, ref int loopIndex)
        {
            string loopNum = BistAction.GetLoopNum(row);
            if (!string.IsNullOrEmpty(loopNum))
            {
                string loopSyntax = _loopSyntax[loopIndex];
                var flowRow = new FlowRow { ColumnA = row.Label, Opcode = OpCode.AssignInteger, Parameter = $"{loopSyntax} -1" };
                flowRows.Add(AddCondition(row, flowRow));
                flowRow = new FlowRow { ColumnA = row.Label, Opcode = OpCode.For, Parameter = $"{loopSyntax}=0;{loopSyntax}<{loopNum};{loopSyntax}++" };
                flowRows.Add(AddCondition(row, flowRow));
                flowRows.AddRange(controlFlagInitRows);
                flowRow = new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = "Control_" + row.Label };
                flowRows.Add(flowRow);
                flowRow = new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = "Control_" + row.PassBranch };
                flowRows.Add(flowRow);
                loopIndex++;
                loopSyntaxStack.Push(loopSyntax);
            }
        }

        private void AddLoopEndRows(BistProdFlowRow row, FlowRows flowRows, Stack<string> loopSyntaxStack)
        {
            string loopSyntax = loopSyntaxStack.Pop();
            flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.Next, DeviceCondition = OpCode.FlagTrue, DeviceName = "Control_" + row.Label });
            flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.AssignInteger, Parameter = $"{loopSyntax} -1" });
            flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = "Control_" + row.PassBranch, DeviceCondition = OpCode.FlagTrue, DeviceName = "Control_" + row.Label });
        }

        private void AddGetRows(BistProdFlowRow row, FlowRows flowRows)
        {
            flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.If, Parameter = BistAction.GetActionParameter(row) });
            flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = GetBranchByLabel(row.PassBranch) }));
            if (CheckBranchType(row.PassBranch) == BistActionType.Fail)
            {
                flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = "F_" + row.Label });
            }

            flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.Else });
            flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = GetBranchByLabel(row.FailBranch) }));
            if (CheckBranchType(row.FailBranch) == BistActionType.Fail)
            {
                flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = GetBranchByLabel(row.FailBranch) });
            }

            flowRows.Add(new FlowRow { ColumnA = row.Label, Opcode = OpCode.EndIf });
        }

        private void AddSetRows(BistProdFlowRow row, FlowRows flowRows)
        {
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(new List<string> { BistAction.SetActionParameter(row) }, "Mbist", FolderStructure.DirMain);
            flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Parameter = BistAction.SetActionParameter(row), Opcode = BistAction.SetActionValue(row).Equals(0) ? OpCode.FlagFalse : OpCode.FlagTrue }));
            flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Opcode = OpCode.FlagTrue, Parameter = GetBranchByLabel(row.PassBranch) }));
        }

        private void AddSetDefaultRows(BistProdFlowRow row, FlowRows flowRows)
        {
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(new List<string> { BistAction.SetDefaultActionParameter(row) }, "Mbist", FolderStructure.DirMain);
            flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Parameter = BistAction.SetDefaultActionParameter(row), Opcode = OpCode.FlagTrue }));
            flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Parameter = GetBranchByLabel(row.PassBranch), Opcode = OpCode.FlagTrue }));
        }

        private void AddFailRows(BistProdFlowRow row, FlowRows flowRows, List<BinTableRow> needBinoutList)
        {
            foreach (string flag in GetAllFailLabelByBranch(row.Label))
            {
                List<string> binNameSeg = flag.Split('_').ToList();
                binNameSeg[0] = binNameSeg[0].Replace("F", "Bin");
                string binName = string.Join("_", binNameSeg);
                flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Opcode = OpCode.BinTable, Parameter = binName }));
                needBinoutList.Add(new BinTableRow { Name = binName, ItemList = flag, Items = new List<string> { "T" }, Op = "AND" });
            }

            foreach (string notestr in row.Note.Split(';'))
            {
                if (notestr.StartsWith("BinItem:"))
                {
                    IEnumerable<string> binitem = notestr.Replace("BinItem:", "").Split(',').Select(s => s);
                    string binname = $"Bin{row.Label}".Replace("_99900", "");
                    flowRows.Add(AddCondition(row, new FlowRow { ColumnA = row.Label, Opcode = OpCode.BinTable, Parameter = binname }));
                    needBinoutList.Add(new BinTableRow { Name = binname, ItemList = string.Join(",", binitem), Items = Enumerable.Repeat("T", binitem.Count()).ToList(), Op = "AND" });
                    break;
                }
            }
        }

        private void ApplyDeviceCondition(BistProdFlowRow row, FlowRows flowRows, Dictionary<string, string> devConRf)
        {
            if (devConRf.ContainsKey(row.Label))
            {
                devConRf.Remove(row.Label);
            }

            if (devConRf.Any())
            {
                string dCflag = devConRf.First().Value;
                flowRows.ForEach(fr => fr.DeviceCondition = BistConst.ConFlagTrue);
                flowRows.ForEach(fr => fr.DeviceName = dCflag);
            }
            if (LocalSpecs.Options.Device == EnumDevice.RF &&
                CheckBranchType(row.PassBranch) == BistActionType.RunPattern &&
                CheckBranchType(row.FailBranch) != BistActionType.Fail &&
                row.PassBranch != row.FailBranch)
            {
                flowRows.First().PassAction = "Control_" + row.Label;
                devConRf.Add(row.FailBranch, "Control_" + row.Label);
            }
        }

        private void PassBranchCheck()
        {
            BistProdFlowRow lastRow = _prodFlowSheet.Rows.LastOrDefault();
            BistProdFlowRow passRow = _prodFlowSheet.Rows.FirstOrDefault(x => BistAction.GetActionType(x) == BistActionType.Pass);
            if (BistAction.GetActionType(lastRow) != BistActionType.Pass)
            {
                string message = passRow.Label + " should be last row in " + _prodFlowSheet.MbistSheet.SheetName + "!!!";
                ErrorReportManager.AddError(MbistErrorType.E_Business_01, _prodFlowSheet.MbistSheet.SheetName, passRow.RowNum, 1, passRow.Label + " should be last row in " + _prodFlowSheet.MbistSheet.SheetName + "!!!", new string[] { passRow.Label, _prodFlowSheet.MbistSheet.SheetName });
                Response.Report(message, EnumMessageLevel.Error, 0);
            }
        }

        private void SheetNameInit()
        {
            string job = _prodFlowSheet.MbistSheet.SheetName.Contains("@") ?
               _prodFlowSheet.MbistSheet.SheetName.Substring(_prodFlowSheet.MbistSheet.SheetName.IndexOf('@') + 1).Replace(",", "_") : "";

            string sourceSheetName = _prodFlowSheet.MbistSheet.SheetName;
            string lStrSheetName = "Flow_" + sourceSheetName;
            if (_isCof)
            {
                lStrSheetName += "_COF";
            }

            _subFlowSheet = new SubFlowSheet(lStrSheetName, sourceSheetName);
            if (string.IsNullOrEmpty(job))
            {
                _subFlowSheet.JobNames = new List<string> { job };
            }
        }

        private void GetPrimeFlags()
        {
            foreach (BistProdFlowRow dataRow in _prodFlowSheet.Rows)
            {
                if (dataRow.BurstPatterns.Any())
                {
                    foreach (string pattern in dataRow.BurstPatterns)
                    {
                        if (_bistNaming.JudgeRepairPattern(pattern, _module))
                        {
                            _primeFlagList.Add(dataRow.FailFlag);
                        }
                    }
                }
                else if (_bistNaming.JudgeRepairPattern(dataRow.Pattern, _module))
                {
                    _primeFlagList.Add(dataRow.FailFlag);
                }
            }
        }

        private void GetPrimeFlagsByDomain()
        {
            string currentDomain = "";
            var currentFlagList = new List<string>();
            foreach (BistProdFlowRow dataRow in _prodFlowSheet.Rows)
            {
                if (BistAction.GetActionType(dataRow) == BistActionType.DomainStart)
                {
                    currentDomain = BistAction.GetDomainByAction(dataRow);
                    continue;
                }
                if (BistAction.GetActionType(dataRow) == BistActionType.DomainEnd)
                {
                    if (!string.IsNullOrEmpty(currentDomain) && currentFlagList.Any())
                    {
                        _domainPrimeFlag.Add(currentDomain, currentFlagList);
                    }

                    currentDomain = "";
                    currentFlagList = new List<string>();
                    continue;
                }
                if (dataRow.BurstPatterns.Any())
                {
                    foreach (string pattern in dataRow.BurstPatterns)
                    {
                        if (_bistNaming.JudgeRepairPattern(pattern, _module))
                        {
                            currentFlagList.Add(dataRow.FailFlag);
                        }
                    }
                }
                else if (_bistNaming.JudgeRepairPattern(dataRow.Pattern, _module))
                {
                    currentFlagList.Add(dataRow.FailFlag);
                }
            }
        }

        private void CheckPatternType(BistProdFlowRow row, out bool isEfp, out bool isRbox, out bool isRepair)
        {
            isEfp = false;
            isRbox = false;
            isRepair = false;

            if (row.BurstPatterns.Any())
            {
                foreach (string pattern in row.BurstPatterns)
                {
                    if (!isEfp && !isRbox)
                    {
                        isEfp = _bistNaming.JudgePattern(pattern, "EFP");
                        isRbox = row.Action.Equals("rbox_program", StringComparison.OrdinalIgnoreCase) ||
                            row.Action.Equals("rbox_compare", StringComparison.OrdinalIgnoreCase);
                    }

                    if (!isRepair)
                    {
                        isRepair = _bistNaming.JudgeRepairPattern(pattern, _module);
                    }
                }
            }
            else
            {
                isEfp = _bistNaming.JudgePattern(row.Pattern, "EFP");
                isRbox = row.Action.Equals("rbox_program", StringComparison.OrdinalIgnoreCase) ||
                    row.Action.Equals("rbox_compare", StringComparison.OrdinalIgnoreCase);
                isRepair = _bistNaming.JudgeRepairPattern(row.Pattern, _module);
            }
        }

        private void AddRepairJudgement(SubFlowSheet flow)
        {
            GetPrimeFlags();
            if (_primeFlagList.Count == 0)
            {
                return;
            }
            int index = flow.Rows.FindLastIndex(p => p.Opcode.Equals("return", StringComparison.CurrentCultureIgnoreCase));
            if (index == -1)
            {
                var returnRow = new FlowRow { Opcode = "return" };
                flow.AddRow(returnRow);
                index = flow.Rows.Count - 1;
            }
            _primeFlagList = _primeFlagList.Distinct().ToList();
            string flagName = "F_" + _module + BistConst.ConMbistRepairCheckFinal;
            var row = new FlowRow { Opcode = BistConst.ConOpCodeIf, Parameter = string.Join("||", _primeFlagList) };
            flow.InsertRow(index++, row);

            row = new FlowRow { Opcode = BistConst.ConFlagTrue, Parameter = flagName };
            flow.InsertRow(index++, row);

            row = new FlowRow { Opcode = BistConst.ConOpCodeEndIf };
            flow.InsertRow(index, row);


            IdsRepairSingleton.Instance().AddMbistRepairFlag(_module, flagName);
        }

        private void AddRepairJudgementByDomain(SubFlowSheet flow)
        {
            GetPrimeFlagsByDomain();
            if (!_domainPrimeFlag.Any())
            {
                return;
            }
            foreach (KeyValuePair<string, List<string>> oneDomain in _domainPrimeFlag)
            {
                int index =
                   flow.Rows.FindLastIndex(
                       p => p.Opcode.Equals("return", StringComparison.CurrentCultureIgnoreCase));
                string module = oneDomain.Key;
                List<string> flagList = oneDomain.Value;
                if (index == -1)
                {
                    var returnRow = new FlowRow { Opcode = "return" };
                    flow.AddRow(returnRow);
                    index = flow.Rows.Count - 1;
                }
                flagList = flagList.Distinct().ToList();
                string flagName = "F_" + module + BistConst.ConMbistRepairCheckFinal;
                var row = new FlowRow { Opcode = BistConst.ConOpCodeIf, Parameter = string.Join("||", flagList) };
                flow.InsertRow(index++, row);

                row = new FlowRow { Opcode = BistConst.ConFlagTrue, Parameter = flagName };
                flow.InsertRow(index++, row);


                row = new FlowRow { Opcode = BistConst.ConOpcodeElse };
                flow.InsertRow(index++, row);

                row = new FlowRow { Opcode = BistConst.ConFlagFalse, Parameter = flagName };
                flow.InsertRow(index++, row);

                row = new FlowRow { Opcode = BistConst.ConOpCodeEndIf };
                flow.InsertRow(index, row);
            }
        }

        private HashSet<string> GetAllFailLabelByBranch(string branchName)
        {
            var flagList = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (_passBranchDic.TryGetValue(branchName, out List<BistProdFlowRow> value))
            {
                flagList.AddRange(value.GroupBy(x => x.FailFlag).Select(y => y.Key));
            }

            if (_failBranchDic.TryGetValue(branchName, out List<BistProdFlowRow> value1))
            {
                flagList.AddRange(value1.GroupBy(x => x.FailFlag).Select(y => y.Key));
            }

            return flagList;
        }

        private FlowRow AddCondition(BistProdFlowRow row, FlowRow flowRow)
        {
            if (_labelDic.FirstOrDefault().Key.Equals(row.Label) || LocalSpecs.Options.Device == EnumDevice.RF)
            {
                return flowRow;
            }

            flowRow.DeviceCondition = OpCode.FlagTrue;
            flowRow.DeviceName = "Control_" + row.Label;
            return flowRow;
        }

        private string GetFailFlagByBranch(BistProdFlowRow row)
        {
            var items = new List<string>();
            string flag = "";
            string detailInfo = GetDetailInfomation(row);
            if (string.IsNullOrEmpty(detailInfo))
            {
                detailInfo = row.Label;
            }

            items.Add(_module);
            items.Add(_conMbist);
            items.Add(detailInfo);
            string level = BistNaming.GetVoltageType(row.Voltage);
            items.Add(level);
            flag = string.Join("_", items).ToUpper();
            if (_binTableType == MbistBinTableType.Single)
            {
                while (_usedFlagList.Exists(x => x.Equals(flag)))
                {
                    List<string> allsegments = flag.Split('_').ToList();
                    int.TryParse(allsegments[allsegments.Count - 2], out int oriIndex);
                    if (oriIndex == 0)
                    {
                        flag = flag.Replace($"{detailInfo}", $"{detailInfo}_1");
                    }
                    else
                    {
                        flag = flag.Replace($"{detailInfo}_{oriIndex}", $"{detailInfo}_{oriIndex + 1}");
                    }
                }
            }
            if (!_usedFlagList.Contains(flag))
            {
                _usedFlagList.Add(flag);
            }

            return flag;
        }

        private string GetDetailInfomation(BistProdFlowRow row)
        {
            return _binTableType == MbistBinTableType.Burst ? row.Label : GetDetailInfoFromPattern(row);
        }

        private string GetDetailInfoFromPattern(BistProdFlowRow row)
        {
            string pattern = row.Pattern;
            List<string> segments = pattern.Split('_').ToList();
            var selIdx = new List<int> { 5, 6, 8, 11 };
            var selItems = new List<string>();
            for (int idx = 0; idx < segments.Count; idx++)
            {
                if (!selIdx.Exists(x => x.Equals(idx)))
                {
                    continue;
                }

                selItems.Add(segments[idx]);
            }
            return string.Join("_", selItems);
        }

        private FlowRow AddAction(BistProdFlowRow row, FlowRow flowRow, bool isRepair = false)
        {
            if (row.Note.Contains("AddFailFlag:"))
            {
                foreach (string notestr in row.Note.Split(';'))
                {
                    if (notestr.StartsWith("AddFailFlag:"))
                    {
                        flowRow.FailAction = notestr.Split(':')[1];
                        row.FailFlag = notestr.Split(':')[1];
                        break;
                    }
                }
            }
            else if (LocalSpecs.Options.Device != EnumDevice.RF)
            {
                var failActionFlags = new List<string> { GetBranchByLabel(row.FailBranch) };
                if (CheckBranchType(row.FailBranch) == BistActionType.Fail || isRepair)
                {
                    failActionFlags.Add(row.FailFlag);
                }
                if (row.FailBranch.ToUpper() != "NEXT")
                {
                    flowRow.FailAction = string.Join(",", failActionFlags.ToArray());
                }

                if (row.PassBranch.ToUpper() != "NEXT")
                {
                    flowRow.PassAction = GetBranchByLabel(row.PassBranch);
                }
            }
            else
            {
                flowRow.FailAction = $"F_{row.Label}";
                row.FailFlag = $"F_{row.Label}";
            }

            return flowRow;
        }

        internal string GetBranchByLabel(string label)
        {
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                return label;
            }

            string result = label;
            if (_labelDic.TryGetValue(label, out List<BistProdFlowRow> value))
            {
                BistProdFlowRow nextBranch = value.FirstOrDefault();
                if (nextBranch == null)
                {
                    return "Control_" + result;
                }

                if (BistAction.GetActionType(nextBranch) == BistActionType.RetentionVoltDrop)
                {
                    result = nextBranch.PassBranch;
                }
            }
            return "Control_" + result;
        }

        internal BistActionType CheckBranchType(string branchName)
        {
            BistActionType type = BistActionType.Null;
            if (_prodFlowSheet.TryGetSheetRow(branchName, out BistProdFlowRow targetBranch))
            {
                return BistAction.GetActionType(targetBranch);
            }
            return type;
        }

        private FlowRow GenerateRetention(BistProdFlowRow priousRow, BistProdFlowRow row)
        {
            string waitTime = BistAction.GetRetentionTime(row);
            string step = BistAction.GetRetentionRampStep(priousRow);
            string sheetName = row.IsMultipleMbistScghSheet ? row.SheetName : "";
            string isWaitTimeOnly = _bistNaming.GetWaitTimeOnly(BistAction.GetActionType(priousRow) != BistActionType.RetentionVoltDrop);
            string lStrParameter = _bistNaming.CreateRetentionTestNameNew(_module, row.Voltage.Split(',', ' ').First(), sheetName, waitTime, step, isWaitTimeOnly);

            return AddCondition(row, new FlowRow { ColumnA = row.Label, Opcode = OpCode.Test, Parameter = lStrParameter });
        }
    }
}
