using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.DividerManager;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class HardIpBlockInsGenerator : BlockInstanceGenerator
    {
        private const string TestSequence = "TESTSEQUENCE";
        private const string CusStrDigcapdata = "CUS_STR_DIGCAPDATA";
        private const string DigsrcEquation = "DIGSRC_EQUATION";
        private const string DigsrcAssignment = "DIGSRC_ASSIGNMENT";
        private const string CalcEqn = "CALC_EQN";
        private const string MeasiRange = "MEASI_RANGE";
        private const string ForcevVal = "FORCEV_VAL";
        private const string ForceiVal = "FORCEI_VAL";
        private const string CusStrMainprogram = "CUS_STR_MAINPROGRAM";

        public HardIpBlockInsGenerator(HardIpInputData hardIpInputData, string sheetName, HardIpSheet hardIpSheet)
            : base(hardIpInputData, sheetName, hardIpSheet)
        {
            InstanceRowGenerator = new HardIpInsRowGenerator(hardIpInputData, hardIpSheet, sheetName);
        }

        /// <summary>
        /// Generate instance rows for one test plan HardIp sheet
        /// </summary>
        public override List<InstanceSheet> GenBlockInsRows()
        {
            List<InstanceSheet> instanceSheetList = new List<InstanceSheet>();
            List<HardIpPattern> patLstToGen = DividerMain.DivideInstancePattern(HardIpInputData, HardIpSheet.Rows).ToList();

            instanceSheetList.AddRange(GenBlockInsRowsByBlock(patLstToGen, HardIpConstData.LabelNv));
            instanceSheetList.AddRange(GenBlockInsRowsByBlock(patLstToGen, HardIpConstData.LabelLv));
            instanceSheetList.AddRange(GenBlockInsRowsByBlock(patLstToGen, HardIpConstData.LabelHv));

            return instanceSheetList;
        }

        internal List<InstanceSheet> GenBlockInsRowsByBlock(List<HardIpPattern> patternList, string labelVoltage)
        {
            var instanceSheetList = new List<InstanceSheet>();
            string blockName = CommonGenerator.GetHardipSheetName(SheetName).ToUpper();

            InstanceSheet block;
            InstanceSheet wireless;
            InstanceSheet splitCzFlow;

            string sheetName = "Inst_H_" + string.Join("_", blockName.Split('_').Skip(1));
            block = new InstanceSheet(sheetName, SheetName);
            wireless = new InstanceSheet(sheetName, SheetName);
            splitCzFlow = new InstanceSheet(sheetName + "_CZ", SheetName);

            var charInsSheet = new InstanceSheet(HardIpConstData.PrefixInsSheetByVoltage + HardIpConstData.LabelChar); //TestInst_HARDIP_Char
            string mtdPatName = "";
            var lastInsRow = new InstanceRow();
            int mtdPatNum = 0;

            Dictionary<string, string> mtdTestName = new Dictionary<string, string>();

            foreach (HardIpPattern pattern in patternList)
            {
                InstanceRowGenerator.LabelVoltage = labelVoltage;
                InstanceRowGenerator.Pat = pattern;
                List<InstanceRow> insRowList = InstanceRowGenerator.GenInsRows();
                if (pattern.Pattern.IsMultiTimeDomain())
                {
                    if (insRowList[0].VbtType == "VBT")
                    {
                        if (mtdPatName == "")
                        {
                            mtdPatName = pattern.Pattern.RealPatternName;
                            lastInsRow = insRowList[0];
                            mtdPatNum++;
                            continue;
                        }

                        if (mtdPatName.ContainsIgnoreCase(pattern.Pattern.TestPlanPatternName))
                        {
                            insRowList[0].MergeVbtMtdArgs(lastInsRow);

                            int registerIndex = HardIpSheet.PlanHeaderIdx["registerIndex"];

                            ProcessMtdArgs(insRowList[0], pattern, blockName, registerIndex);

                            if (pattern.Pattern.RealPatternName.Split('#').Length == mtdPatNum + 1)
                            {
                                lastInsRow = new InstanceRow();
                                mtdPatNum = 0;
                                mtdPatName = "";
                            }
                            else if (mtdPatName.Contains(pattern.Pattern.TestPlanPatternName))
                            {
                                lastInsRow = insRowList[0];
                                mtdPatNum++;
                                continue;
                            }
                        }
                    }
                    else
                    {
                        if (pattern.MiscInfoDict.ContainsKey("Ref_SubBlock") && pattern.MiscInfoDict["Ref_SubBlock"].Contains("#"))
                        {
                            List<string> insts = new List<string>();
                            foreach (string refsubblk in pattern.MiscInfoDict["Ref_SubBlock"].Split('#'))
                            {
                                if (mtdTestName.TryGetValue(refsubblk, out string value))
                                {
                                    insts.Add(value);
                                }
                            }
                            insRowList[0].SetArgument("instanceNameList", string.Join(",", insts));
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                else if (!mtdTestName.ContainsKey(pattern.SubBlock))
                {
                    mtdTestName.Add(pattern.SubBlock, insRowList.First().TestName);
                }

                if (pattern.SkipList.Contains(labelVoltage) || pattern.SkipDotNet)
                {
                    continue;
                }

                foreach (InstanceRow insRow in insRowList)
                {
                    AddInsRowToSheet(insRow, pattern, labelVoltage, block, wireless, charInsSheet);
                }
            }

            if (block.Rows.Count != 0)
            {
                instanceSheetList.Add(block);
            }

            if (wireless.Rows.Count != 0)
            {
                instanceSheetList.Add(wireless);
            }

            if (charInsSheet.Rows.Count != 0)
            {
                instanceSheetList.Add(charInsSheet);
            }

            if (splitCzFlow.Rows.Count != 0)
            {
                instanceSheetList.Add(splitCzFlow);
            }

            return instanceSheetList;
        }

        private void ProcessMtdArgs(InstanceRow insRow, HardIpPattern pattern, string blockName, int registerIndex)
        {
            List<string> parameters = insRow.ArgList.Split(',').ToList();
            for (int i = 0; i < parameters.Count; i++)
            {
                string infoParameter = insRow.Args[i];
                string type = parameters[i];
                // For all args which start with "+"
                if (!string.IsNullOrEmpty(infoParameter) && infoParameter.StartsWith("+"))
                {
                    insRow.Args[i] = "'" + insRow.Args[i];
                    infoParameter = insRow.Args[i];
                }

                if (infoParameter.Length >= 6000)
                {
                    string[] infopara = infoParameter.Split('#');
                    List<string> regAssNames = new List<string>();
                    for (int j = 0; j < infopara.Length; j++)
                    {
                        var regAssignItem = new HardIpRegAssign();
                        string multiblockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
                        regAssignItem.SubBlockName = multiblockName + "_" + CommonGenerator.GetSubBlockName(pattern.Pattern.GetLastPayload(), pattern.MiscInfo, blockName) + "_" + j;

                        List<List<string>> regAssginList = GetRegAssginList(pattern, registerIndex, type, infopara, j, regAssignItem);

                        //set args
                        regAssNames.Add($"Reg_assign:{regAssignItem.SubBlockName}");
                        if (!HardIpInputData.HardIpRegAssigns.Exists(x => x.SubBlockName.Equals(regAssignItem.SubBlockName, StringComparison.OrdinalIgnoreCase) && x.Type == regAssignItem.Type))
                        {
                            regAssignItem.RegAssignList = regAssginList;
                            HardIpInputData.HardIpRegAssigns.Add(regAssignItem);
                        }

                        //Hint in error Report
                        ErrorReportManager.AddError(
                            HardIpErrorType.W_RegAssignWarning_01,
                            pattern.SheetName,
                            pattern.RowNum,
                            registerIndex,
                            [type]
                        );
                    }

                    insRow.Args[i] = string.Join("#", regAssNames);
                }
            }
        }

        private void AddInsRowToSheet(InstanceRow insRow, HardIpPattern pattern, string labelVoltage, InstanceSheet block, InstanceSheet wireless, InstanceSheet charInsSheet)
        {
            if (Regex.IsMatch(insRow.TestName, "INSREMOV_", RegexOptions.IgnoreCase) && labelVoltage != "NV")
            {
                return;
            }

            insRow.TestName = insRow.TestName.Replace("INSREMOV_", "");
            if (pattern.SheetName.StartsWith(NeededSheets.PrefixWireless) ||
                pattern.SheetName.StartsWith(NeededSheets.PrefixLcd))
            {
                wireless.AddRow(insRow);
            }
            else
            {
                if (InstanceRowGenerator.Pat.ForceCondition.IsShmooInCharInst)
                {
                    charInsSheet.AddRow(insRow);
                }

                if (InstanceRowGenerator.Pat.ForceCondition.IsShmooInProdInst)
                {
                    block.AddRow(insRow);
                }
            }
        }

        internal static List<List<string>> GetRegAssginList(HardIpPattern pattern, int registerIndex, string type, string[] infopara, int j, HardIpRegAssign regAssignItem)
        {
            var regAssginList = new List<List<string>>();
            switch (type.ToUpper())
            {
                case TestSequence:
                    regAssignItem.Type = RegisterAssignType.TestSequence;
                    foreach (string item in infopara[j].Split(',').ToList())
                    {
                        regAssginList.Add(new List<string> { "'" + item });
                    }
                    break;
                case DigsrcEquation: //+
                    regAssignItem.Type = RegisterAssignType.DigSrc_Equation;
                    foreach (string item in infopara[j].Split('+').ToList())
                    {
                        regAssginList.Add(new List<string> { "'" + item });
                    }
                    break;
                case DigsrcAssignment:
                    regAssignItem.Type = RegisterAssignType.DigSrc_Assignment;
                    foreach (string item in infopara[j].Split(';').ToList())
                    {
                        List<string> data = new List<string>();
                        List<string> arr = item.Split('=').ToList();
                        data.Add(arr[0]);
                        if (arr.Count > 1)
                        {
                            data.Add("'" + arr[1]);
                        }

                        regAssginList.Add(data);
                    }
                    break;
                case CusStrDigcapdata:
                    regAssignItem.Type = RegisterAssignType.CUS_Str_DigCapData;
                    foreach (string item in infopara[j].Split(',').ToList())
                    {
                        regAssginList.Add(item.Split(':').ToList());
                    }

                    break;
                case CalcEqn:
                    regAssignItem.Type = RegisterAssignType.Calc_Eqn;
                    foreach (string item in infopara[j].Split(';').ToList())
                    {
                        regAssginList.Add(item.Split(':').ToList());
                    }

                    break;
                case MeasiRange:
                    regAssignItem.Type = RegisterAssignType.MeasI_Range;
                    foreach (string item in infopara[j].Split('+').ToList())
                    {
                        regAssginList.Add(new List<string> { "'" + item });
                    }

                    break;
                case ForcevVal:
                    regAssignItem.Type = RegisterAssignType.ForceV_Val;
                    foreach (string item in infopara[j].Split('|').ToList())
                    {
                        regAssginList.Add(new List<string> { "'" + item });
                    }

                    break;
                case ForceiVal:
                    regAssignItem.Type = RegisterAssignType.ForceI_Val;
                    foreach (string item in infopara[j].Split('|').ToList())
                    {
                        regAssginList.Add(new List<string> { "'" + item });
                    }

                    break;
                case CusStrMainprogram:
                    regAssignItem.Type = RegisterAssignType.CUS_Str_MainProgram;
                    foreach (string item in infopara[j].Split(',').ToList())
                    {
                        regAssginList.Add(new List<string> { "'" + item });
                    }

                    break;
                default:
                    regAssignItem.Type = RegisterAssignType.NoneType;
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_RegAssignError_01,
                        pattern.SheetName,
                        pattern.RowNum,
                        registerIndex,
                        [type]
                    );
                    break;
            }

            return regAssginList;
        }
    }
}
