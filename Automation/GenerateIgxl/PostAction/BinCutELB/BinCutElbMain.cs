using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using LogLib.Utility;

namespace Automation.GenerateIgxl.PostAction.BinCutELB
{
    [ExcludeFromCodeCoverage]
    public class BinCutElbMain
    {
        public void WorkFlow(string trunkPath, string outputFolderNew, string outputFolderOld, InstanceSheet testInstVddBinningSheet, List<SubFlowSheet> flowSheets, ref Dictionary<string, SubFlowSheet> subFlowSheets, ref Dictionary<string, InstanceSheet> instanceSheets, ref List<InstanceSheet> instanceHardIpSheets, List<string> rtosList)
        {
            List<InstanceMappingRow> instanceMappingRows = GetInstanceMappings(testInstVddBinningSheet);
            List<string> instanceSheetNames = IgxlSheetReaderHelpers.GetSheetsByType(trunkPath, EnumSheetType.DTTestInstancesSheet);
            var dic = new Dictionary<string, List<InstanceRow>>();
            foreach (string instanceSheetName in instanceSheetNames)
            {
                var readInstanceSheet = new ReadInstanceSheet();
                InstanceSheet instanceSheet = readInstanceSheet.GetSheet(instanceSheetName);
                if (rtosList != null)
                {
                    foreach (InstanceRow row in instanceSheet.Rows)
                    {
                        foreach (string rtos in rtosList)
                        {
                            List<string> arrRtos = rtos.Split('_').ToList();
                            List<string> arr = row.TestName.Split('_').ToList();
                            if (arrRtos.All(x => arr.Exists(y => y.Equals(x, StringComparison.CurrentCultureIgnoreCase))))
                            {
                                if (!dic.ContainsKey(rtos))
                                {
                                    dic.Add(rtos, new List<InstanceRow> { row });
                                }
                                else
                                {
                                    dic[rtos].Add(row);
                                }
                            }
                        }
                    }
                }

                InstanceSheet oldSheet = instanceSheet.Copy();
                bool flag = false;
                foreach (InstanceRow row in instanceSheet.Rows)
                {
                    string instanceName = row.TestName.ToUpper();
                    if (instanceMappingRows.Exists(x => x.HardipInstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        InstanceMappingRow instanceMappingRow = instanceMappingRows.Find(x => x.HardipInstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase));
                        row.Overlay = GetOverlay(row, instanceMappingRow.HardipInstanceName);
                        flag = true;
                    }
                }

                if (flag)
                {
                    string name = Path.GetFileNameWithoutExtension(instanceSheetName);
                    instanceHardIpSheets.Add(instanceSheet);
                    instanceSheets.Add(Path.Combine(outputFolderNew, name + ".txt"), instanceSheet);
                    instanceSheets.Add(Path.Combine(outputFolderOld, name + ".txt"), oldSheet);
                }
            }

            #region Add Rtos insance
            bool isChange = false;
            InstanceSheet oldTestInstVddBinningSheet = testInstVddBinningSheet.Copy();
            var instanceRowDic = new Dictionary<string, InstanceRow>();
            foreach (InstanceRow row in testInstVddBinningSheet.Rows)
            {
                foreach (KeyValuePair<string, List<InstanceRow>> item in dic)
                {
                    if (row.TestName.ToUpper().Contains(item.Key.ToUpper()))
                    {
                        if (item.Value.Count > 1)
                        {
                            ErrorMessageBox.Show("Find out more than one instance by binCut RTOS condition ,please check RTOS instance sheet !!! ");
                        }

                        string testName = "";
                        string last = row.TestName.Split('_').ToList().Last();
                        InstanceRow instance = item.Value.First();
                        List<string> arr = instance.TestName.Split('_').ToList();
                        arr.RemoveAt(arr.Count - 1);
                        if (last.Equals("HBV", StringComparison.CurrentCultureIgnoreCase))
                        {
                            testName = string.Join("_", arr) + "_HV";
                        }
                        else if (last.Equals("BV", StringComparison.CurrentCultureIgnoreCase))
                        {
                            testName = string.Join("_", arr) + "_LV";
                        }

                        instance.TestName = testName;
                        instance.Overlay = "Overlay_BV_" + instance.TestName;
                        instanceRowDic.Add(row.TestName, instance);
                        row.VbtName = "GradeSearch_HVCC_CallInstance_VT";
                        row.SetArgument("inst_CallInstance", testName);
                        row.SetArgument("Performance_mode", row.Args[0]);
                        isChange = true;
                    }
                }
            }

            if (isChange)
            {
                foreach (KeyValuePair<string, InstanceRow> item in instanceRowDic)
                {
                    testInstVddBinningSheet.AddRow(item.Value);
                }

                string name = Path.GetFileNameWithoutExtension(oldTestInstVddBinningSheet.Name);
                instanceSheets.Add(Path.Combine(outputFolderNew, name + ".txt"), testInstVddBinningSheet);
                instanceSheets.Add(Path.Combine(outputFolderOld, name + ".txt"), oldTestInstVddBinningSheet);
            }
            #endregion

            if (instanceMappingRows.Count == 0)
            {
                return;
            }

            foreach (InstanceMappingRow instanceMappingRow in instanceMappingRows)
            {
                if (!instanceMappingRow.IsFound)
                {
                    ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_03, testInstVddBinningSheet.Name, instanceMappingRow.BinCutRow.RowNum, 0, $"The ELB instance {instanceMappingRow.HardipInstanceName} can not be found !!!", new string[] { instanceMappingRow.HardipInstanceName });
                }
            }

            #region Get flow sheet
            List<string> flowSheetFiles = IgxlSheetReaderHelpers.GetSheetsByType(trunkPath, EnumSheetType.DTFlowtableSheet);
            var hardIpFlowSheets = new List<SubFlowSheet>();
            foreach (string flowsheetFile in flowSheetFiles)
            {
                string name = Path.GetFileNameWithoutExtension(flowsheetFile);
                if (name.StartsWith("Flow_HARDIP_", StringComparison.CurrentCultureIgnoreCase))
                {
                    var readFlowSheet = new ReadFlowSheet();
                    SubFlowSheet subFlowSheet = readFlowSheet.ReadSheet(IgxlSheetReaderHelpers.ConvertTxtToExcelSheet(flowsheetFile));
                    hardIpFlowSheets.Add(subFlowSheet);
                }
            }

            var binCutFlowSheets = flowSheets.Where(x => x.Name.EndsWith("_TD_Mbist_BV", StringComparison.CurrentCultureIgnoreCase) ||
                x.Name.StartsWith("Flow_VddBinning", StringComparison.CurrentCultureIgnoreCase)).ToList();
            #endregion

            Dictionary<string, List<FlowRow>> limitDic = GetHardipLimit(instanceMappingRows, hardIpFlowSheets);

            subFlowSheets = ModifyBinCutFlow(outputFolderNew, outputFolderOld, binCutFlowSheets, instanceMappingRows, limitDic, instanceRowDic);
        }

        public void WorkFlow(List<InstanceSheet> instanceSheets, List<SubFlowSheet> flowSheets, List<string> rtosList)
        {
            List<InstanceSheet> testInstVddBinningSheets = instanceSheets.FindAll(x => x.Name.StartsWith("TestInst_VddBinning", StringComparison.OrdinalIgnoreCase));
            var instanceMappingRows = new List<InstanceMappingRow>();
            foreach (InstanceSheet testInstVddBinningSheet in testInstVddBinningSheets)
            {
                instanceMappingRows.AddRange(GetInstanceMappings(testInstVddBinningSheet));
            }
            var dic = new Dictionary<string, List<InstanceRow>>();
            foreach (InstanceSheet instanceSheet in instanceSheets)
            {
                if (rtosList != null)
                {
                    foreach (InstanceRow row in instanceSheet.Rows)
                    {
                        foreach (string rtos in rtosList)
                        {
                            List<string> arrRtos = rtos.Split('_').ToList();
                            List<string> arr = row.TestName.Split('_').ToList();
                            if (arrRtos.All(x => arr.Exists(y => y.Equals(x, StringComparison.CurrentCultureIgnoreCase))))
                            {
                                if (!dic.ContainsKey(rtos))
                                {
                                    dic.Add(rtos, new List<InstanceRow> { row });
                                }
                                else
                                {
                                    dic[rtos].Add(row);
                                }
                            }
                        }
                    }
                }

                foreach (InstanceRow row in instanceSheet.Rows)
                {
                    string instanceName = row.TestName.ToUpper();
                    if (instanceMappingRows.Exists(x => x.HardipInstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        InstanceMappingRow instanceMappingRow = instanceMappingRows.Find(x => x.HardipInstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase));
                        row.Overlay = GetOverlay(row, instanceMappingRow.HardipInstanceName);
                    }
                }
            }

            #region Add Rtos insance
            var instanceRowDic = new Dictionary<string, InstanceRow>();
            foreach (InstanceSheet testInstVddBinningSheet in testInstVddBinningSheets)
            {
                foreach (InstanceRow row in testInstVddBinningSheet.Rows)
                {
                    foreach (KeyValuePair<string, List<InstanceRow>> item in dic)
                    {
                        if (row.TestName.ToUpper().Contains(item.Key.ToUpper()))
                        {
                            if (item.Value.Count > 1)
                            {
                                ErrorMessageBox.Show("Find out more than one instance by binCut RTOS condition ,please check RTOS instance sheet !!! ");
                            }

                            string testName = "";
                            string last = row.TestName.Split('_').ToList().Last();
                            InstanceRow instance = item.Value.First();
                            List<string> arr = instance.TestName.Split('_').ToList();
                            arr.RemoveAt(arr.Count - 1);
                            if (last.Equals("HBV", StringComparison.CurrentCultureIgnoreCase))
                            {
                                testName = string.Join("_", arr) + "_HV";
                            }
                            else if (last.Equals("BV", StringComparison.CurrentCultureIgnoreCase))
                            {
                                testName = string.Join("_", arr) + "_LV";
                            }

                            instance.TestName = testName;
                            instance.Overlay = "Overlay_BV_" + instance.TestName;
                            instanceRowDic.Add(row.TestName, instance);
                            row.VbtName = "GradeSearch_HVCC_CallInstance_VT";
                            row.SetArgument("inst_CallInstance", testName);
                            row.SetArgument("Performance_mode", row.Args[0]);
                        }
                    }
                }
            }
            #endregion

            if (instanceMappingRows.Count == 0)
            {
                return;
            }

            foreach (InstanceSheet testInstVddBinningSheet in testInstVddBinningSheets)
            {
                foreach (InstanceMappingRow instanceMappingRow in instanceMappingRows)
                {
                    if (!instanceMappingRow.IsFound)
                    {
                        string errMsg = $"The ELB instance {instanceMappingRow.HardipInstanceName} can not be found !!!";
                        ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_04, testInstVddBinningSheet.Name, instanceMappingRow.BinCutRow.RowNum, 0, $"The ELB instance {instanceMappingRow.HardipInstanceName} can not be found !!!", new string[] { instanceMappingRow.HardipInstanceName });
                    }
                }
            }
            #region Get flow sheet

            List<SubFlowSheet> hardIpFlowSheets = flowSheets.FindAll(x => x.Name.StartsWith("Flow_HARDIP_", StringComparison.OrdinalIgnoreCase));
            var binCutFlowSheets = flowSheets.Where(x => x.Name.EndsWith("_TD_Mbist_BV", StringComparison.CurrentCultureIgnoreCase) ||
                x.Name.StartsWith("Flow_VddBinning", StringComparison.CurrentCultureIgnoreCase) ||
                x.Name.StartsWith("Flow_PostBinCut", StringComparison.CurrentCultureIgnoreCase)).ToList();
            #endregion

            Dictionary<string, List<FlowRow>> limitDic = GetHardipLimit(instanceMappingRows, hardIpFlowSheets);

            ModifyBinCutFlow(binCutFlowSheets, instanceMappingRows, limitDic, instanceRowDic);
        }

        private string GetOverlay(InstanceRow row, string hardipInstanceName)
        {
            string overlay = row.Overlay;
            string newOverlay = "Overlay_BV_" + hardipInstanceName;
            if (string.IsNullOrEmpty(overlay))
            {
                overlay = newOverlay;
            }
            else
            {
                List<string> arr = overlay.Split(',').ToList();
                if (!arr.Exists(x => x.Equals(newOverlay, StringComparison.CurrentCultureIgnoreCase)))
                {
                    arr.Add(newOverlay);
                    overlay = string.Join(",", arr);
                }
            }
            return overlay;
        }

        private Dictionary<string, SubFlowSheet> ModifyBinCutFlow(string outputFolderNew, string outputFolderOld, List<SubFlowSheet> binCutSheets, List<InstanceMappingRow> instanceMappingRows, Dictionary<string, List<FlowRow>> limitDic, Dictionary<string, InstanceRow> instanceRowDic = null)
        {
            var subFlowSheets = new Dictionary<string, SubFlowSheet>();
            foreach (SubFlowSheet binCutSheet in binCutSheets)
            {
                SubFlowSheet oldSheet = binCutSheet.Copy();
                bool flag = false;
                for (int index = 0; index < binCutSheet.Rows.Count; index++)
                {
                    FlowRow row = binCutSheet.Rows[index];
                    if (row.Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase) ||
                        row.Opcode.Equals("Nop", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (row.Parameter.ToUpper().Contains("_CALLINST"))
                        {
                            if (string.IsNullOrEmpty(row.FailAction))
                            {
                                row.FailAction = "F_BV_CALLINST";
                            }
                            else
                            {
                                List<string> arr = row.FailAction.Split(',').ToList();
                                if (!arr.Exists(x => x.Equals("F_BV_CALLINST", StringComparison.CurrentCultureIgnoreCase)))
                                {
                                    row.FailAction += ",F_BV_CALLINST";
                                }
                            }
                        }
                        if (instanceMappingRows.Exists(x => x.BinCutInstanceNames.Exists(y => y.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase))))
                        {
                            InstanceMappingRow hardipInstance = instanceMappingRows.Find(x => x.BinCutInstanceNames.Exists(y => y.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase)));
                            if (limitDic.TryGetValue(hardipInstance.HardipInstanceName, out List<FlowRow> value))
                            {
                                var copyRows = new List<FlowRow>();
                                foreach (FlowRow limitRow in value)
                                {
                                    if (limitRow.Opcode.Equals(OpCode.Test, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        row.Comment = !string.IsNullOrEmpty(limitRow.Comment) ? limitRow.Comment : "";
                                        row.Comment1 = !string.IsNullOrEmpty(limitRow.Comment1) ? limitRow.Comment1 : "";
                                    }
                                    else
                                    {
                                        FlowRow copyRow = limitRow.Copy();
                                        copyRow.Parameter = row.Parameter;
                                        copyRow.FailAction = "F_BV_CALLINST";
                                        copyRows.Add(copyRow);
                                    }
                                }
                                binCutSheet.RemoveLimitRows(index);
                                binCutSheet.Rows.InsertRange(index + 1, copyRows);
                                flag = true;
                            }
                        }
                    }
                }

                #region RTOS

                if (instanceRowDic != null)
                {
                    foreach (KeyValuePair<string, InstanceRow> item in instanceRowDic)
                    {
                        if (binCutSheet.Rows.Exists(x => x.Parameter.Equals(item.Key, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var flowRow = new FlowRow { Enable = "ForCallInst", Opcode = OpCode.Test, Parameter = item.Value.TestName };
                            binCutSheet.Rows.Insert(binCutSheet.Rows.Count, flowRow);
                            flag = true;
                        }
                    }
                }

                #endregion

                if (flag)
                {
                    subFlowSheets.Add(Path.Combine(outputFolderNew, binCutSheet.Name + ".txt"), binCutSheet);
                    subFlowSheets.Add(Path.Combine(outputFolderOld, binCutSheet.Name + ".txt"), oldSheet);
                }
            }
            return subFlowSheets;
        }

        private void ModifyBinCutFlow(List<SubFlowSheet> binCutSheets,
            List<InstanceMappingRow> instanceMappingRows, Dictionary<string, List<FlowRow>> limitDic, Dictionary<string, InstanceRow> instanceRowDic = null)
        {
            foreach (SubFlowSheet binCutSheet in binCutSheets)
            {
                for (int index = 0; index < binCutSheet.Rows.Count; index++)
                {
                    FlowRow row = binCutSheet.Rows[index];
                    if (row.Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase) ||
                        row.Opcode.Equals("Nop", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (row.Parameter.ToUpper().Contains("_CALLINST"))
                        {
                            if (string.IsNullOrEmpty(row.FailAction))
                            {
                                row.FailAction = "F_BV_CALLINST";
                            }
                            else
                            {
                                List<string> arr = row.FailAction.Split(',').ToList();
                                if (!arr.Exists(x => x.Equals("F_BV_CALLINST", StringComparison.CurrentCultureIgnoreCase)))
                                {
                                    row.FailAction += ",F_BV_CALLINST";
                                }
                            }
                        }
                        if (instanceMappingRows.Exists(x => x.BinCutInstanceNames.Exists(y => y.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase))))
                        {
                            InstanceMappingRow hardipInstance = instanceMappingRows.Find(x => x.BinCutInstanceNames.Exists(y => y.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase)));
                            if (limitDic.TryGetValue(hardipInstance.HardipInstanceName, out List<FlowRow> value))
                            {
                                var copyRows = new List<FlowRow>();
                                foreach (FlowRow limitRow in value)
                                {
                                    if (limitRow.Opcode.Equals(OpCode.Test, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        row.Comment = !string.IsNullOrEmpty(limitRow.Comment) ? limitRow.Comment : "";
                                        row.Comment1 = !string.IsNullOrEmpty(limitRow.Comment1) ? limitRow.Comment1 : "";
                                    }
                                    else
                                    {
                                        FlowRow copyRow = limitRow.Copy();
                                        copyRow.FailAction = "F_BV_CALLINST";
                                        copyRows.Add(copyRow);
                                    }
                                }
                                binCutSheet.RemoveLimitRows(index);
                                binCutSheet.Rows.InsertRange(index + 1, copyRows);
                            }
                        }
                    }
                }

                #region RTOS

                if (instanceRowDic != null)
                {
                    foreach (KeyValuePair<string, InstanceRow> item in instanceRowDic)
                    {
                        if (binCutSheet.Rows.Exists(x => x.Parameter.Equals(item.Key, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var flowRow = new FlowRow { Enable = "ForCallInst", Opcode = OpCode.Test, Parameter = item.Value.TestName };
                            binCutSheet.Rows.Insert(binCutSheet.Rows.Count, flowRow);
                        }
                    }
                }

                #endregion
            }
        }

        private Dictionary<string, List<FlowRow>> GetHardipLimit(List<InstanceMappingRow> instanceMappingRows, List<SubFlowSheet> subFlowSheets)
        {
            var limitDic = new Dictionary<string, List<FlowRow>>();
            var instanceNames = instanceMappingRows.Select(x => x.HardipInstanceName).Distinct().ToList();
            foreach (string instanceName in instanceNames)
            {
                foreach (SubFlowSheet subFlowSheet in subFlowSheets)
                {
                    if (subFlowSheet.Rows.Exists(x => x.Parameter.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        int index = subFlowSheet.Rows.FindIndex(x => x.Parameter.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase) &&
                                                                   (x.Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase) ||
                                                                    x.Opcode.Equals("Nop", StringComparison.CurrentCultureIgnoreCase)));
                        if (index != -1)
                        {
                            if (!limitDic.ContainsKey(instanceName))
                            {
                                List<FlowRow> flowRows = subFlowSheet.GetTestAndLimitRows(index);
                                flowRows.ForEach(x => x.ColumnA = "");
                                flowRows.ForEach(x => x.FailAction = "F_BV_CALLINST");
                                limitDic.Add(instanceName, flowRows);
                            }
                            else
                            {
                                ErrorReportManager.AddError(BasicErrorType.E_FormatWarning_05, subFlowSheet.Name, subFlowSheet.Rows[index].RowNum, 0, $"The limit of ELB instance are duplicated -{instanceName} !!!", new string[] { instanceName });
                            }
                        }
                    }
                }
            }
            return limitDic;
        }

        private List<InstanceMappingRow> GetInstanceMappings(InstanceSheet instanceSheet)
        {
            var instanceRows = new List<InstanceRow>();
            foreach (InstanceRow row in instanceSheet.Rows)
            {
                if (row.TestName.Contains("CallInst"))
                {

                }
                if (row.VbtName.Equals("GradeSearch_CallInstance_VT", StringComparison.CurrentCultureIgnoreCase) ||
                    row.VbtName.Equals("GradeSearch_HVCC_CallInstance_VT", StringComparison.CurrentCultureIgnoreCase) ||
                    !string.IsNullOrEmpty(row.GetArgument("callInstanceName")))
                {
                    instanceRows.Add(row);
                }
            }

            var instanceMappings = new List<InstanceMappingRow>();
            foreach (InstanceRow instanceRow in instanceRows)
            {
                if (instanceRow.Args.Count > 5)
                {
                    string hardipInstanceName = instanceRow.GetArgument("callInstanceName");
                    if (string.IsNullOrEmpty(hardipInstanceName))
                    {
                        hardipInstanceName = instanceRow.GetArgument("inst_CallInstance");
                    }

                    if (string.IsNullOrEmpty(hardipInstanceName))
                    {
                        continue;
                    }

                    if (instanceMappings.Exists(x => x.HardipInstanceName.Equals(hardipInstanceName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        InstanceMappingRow instanceMappingRow = instanceMappings.Find(x => x.HardipInstanceName.Equals(hardipInstanceName, StringComparison.CurrentCultureIgnoreCase));
                        instanceMappingRow.BinCutInstanceNames.Add(instanceRow.TestName);
                        instanceMappingRow.BinCutRow = instanceRow;
                        instanceMappingRow.IsFound = true;
                    }
                    else
                    {
                        var instanceMappingRow = new InstanceMappingRow { HardipInstanceName = hardipInstanceName };
                        instanceMappingRow.BinCutInstanceNames.Add(instanceRow.TestName);
                        instanceMappingRow.BinCutRow = instanceRow;
                        instanceMappingRow.IsFound = true;
                        instanceMappings.Add(instanceMappingRow);
                    }
                }
            }
            return instanceMappings;
        }
    }
}
