using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

using CommonLib.Extension;
using CommonLib.Utility;

using IgxlLib.Const;
using IgxlLib.Enums;
using IgxlLib.IGDataXML.IGXLSheetsVersion;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;

using OfficeOpenXml;

namespace IgxlLib.IgxlSheets
{
    public partial class SubFlowSheet : IgxlSheet<FlowRow>
    {
        public class IfCondition
        {
            public string If = string.Empty;
            public int IfIndex = -1;
            public string ElseIf = string.Empty;
            public int ElseIfIndex = -1;
            public string Else = string.Empty;
            public int ElseIndex = -1;
            public string EndIf = string.Empty;
            public int EndIfIndex = -1;
        }

        [GeneratedRegex("efuse_", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex EfusePrefixRegex();

        public override string SheetType => "DTFlowtableSheet";
        public override string IgxlSheetName { get; } = IgxlSheetNames.FlowTable;

        public List<string> JobNames { get; set; } = [];
        public EnumGroupInMainFlow GroupNameInMainFlow = EnumGroupInMainFlow.None;
        public bool HasConcurrent = false;
        public string SplitFromSheet = string.Empty;

        public SubFlowSheet(ExcelWorksheet excelWorksheet) : base(excelWorksheet)
        {
            IgxlSheetName = IgxlSheetNames.FlowTable;
        }

        public SubFlowSheet(string sheetName, string sourceSheet = "") : base(sheetName)
        {
            SourceInfo.Name = sourceSheet;
        }

        public SubFlowSheet(SubFlowSheet subFlowSheet) : base(subFlowSheet)
        {
            if (subFlowSheet == null)
            {
                return;
            }

            JobNames = [.. subFlowSheet.JobNames];
            GroupNameInMainFlow = subFlowSheet.GroupNameInMainFlow;
            HasConcurrent = subFlowSheet.HasConcurrent;
            SplitFromSheet = subFlowSheet.SplitFromSheet;
            Rows = [.. subFlowSheet.Rows.Select(x => x.Copy())];
        }

        public SubFlowSheet Copy()
        {
            return new SubFlowSheet(this);
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
            {
                version = "3.0";
            }

            WriteByOneVersion(fileName, WriteContent(version));
        }

        public StringBuilder? WriteContent(string version = "")
        {
            Dictionary<string, Dictionary<string, SheetInfo>> sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.TryGetValue(IgxlSheetName, out Dictionary<string, SheetInfo>? dic))
            {
                if (dic.TryGetValue(version, out SheetInfo? igxlSheetsVersion1))
                {
                    return WriteSheet(version, igxlSheetsVersion1);
                }
                if (dic.TryGetValue("3.0", out SheetInfo? igxlSheetsVersion))
                {
                    return WriteSheet(version, igxlSheetsVersion);
                }
            }
            return null;
        }

        public void AddReturnRow()
        {
            var flowRow = new FlowRow
            {
                Opcode = OpCode.Return
            };
            Rows.Add(flowRow);
        }

        public void AddStartRows(string sheetName = "")
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                sheetName = Name;
            }

            string name = sheetName.ReplaceStartsWith("Flow_", "");
            //Header
            AddHeaderRow(name);
            //Print
            AddPrintStartRow(name);
        }

        public void AddEndRows(string sheetName = "")
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                sheetName = Name;
            }

            string name = sheetName.ReplaceStartsWith("Flow_", "");
            //Print
            AddPrintEndRow(name);
            //Footer
            AddFooterRow(name);
            //Return
            AddReturnRow();
        }

        public void AddHeaderRow(string name, string enable = "", string rowIdx = "")
        {
            var flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = name + "_Header_1", Enable = enable };
            if (string.IsNullOrEmpty(rowIdx))
            {
                Rows.Add(flowRow);
            }
            else
            {
                InsertRow(int.Parse(rowIdx), flowRow);
            }
        }

        public void AddFooterRow(string name, string enable = "", string rowIdx = "")
        {
            var flowRow = new FlowRow { Opcode = OpCode.Test, Parameter = name + "_Footer_1", Enable = enable };
            if (string.IsNullOrEmpty(rowIdx))
            {
                Rows.Add(flowRow);
            }
            else
            {
                InsertRow(int.Parse(rowIdx), flowRow);
            }
        }

        public void AddPrintStartRow(string sheetName)
        {
            var flowRow = new FlowRow
            {
                Opcode = OpCode.Print,
                Parameter = "\"Flow " + sheetName + " Start\""
            };
            Rows.Add(flowRow);
        }

        public void AddPrintEndRow(string sheetName)
        {
            var flowRow = new FlowRow
            {
                Opcode = OpCode.Print,
                Parameter = "\"Flow " + sheetName + " Stop\""
            };
            Rows.Add(flowRow);
        }

        public List<string> GetFailFlags()
        {
            var failFlags = Rows.Where(x => !string.IsNullOrEmpty(x.FailAction))
                    .Select(x => x.FailAction.Split(',')).SelectMany(x => x).Distinct(StringExtensions.IgnoreCase).ToList();
            return failFlags;
        }

        public List<string> GetAllInitFalse()
        {
            var falseFlags =
                Rows.Where(x => x.Opcode == OpCode.FlagFalse).Select(x => x.Parameter.Split(',')).SelectMany(x => x).Distinct(StringExtensions.IgnoreCase).ToList();
            return falseFlags;
        }

        public void AddClearFlag(string flag)
        {
            var flowRow = new FlowRow { Opcode = OpCode.FlagClear, Parameter = flag };
            Rows.Add(flowRow);
        }

        public void AddFlagToTrue(string flag, string enableWord)
        {
            var flowRow = new FlowRow { Opcode = OpCode.FlagTrue, Parameter = flag };
            if (!string.IsNullOrEmpty(enableWord))
            {
                flowRow.Enable = enableWord;
            }

            Rows.Add(flowRow);
        }

        public List<FlowRow> GetTestAndLimitRows(int index)
        {
            var flowRows = new List<FlowRow>();
            if (Rows[index].Opcode.EqualsIgnoreCase(OpCode.Test))
            {
                flowRows.Add(Rows[index].Copy());
            }

            for (int i = index + 1; i < Rows.Count; i++)
            {
                FlowRow row = Rows[i];
                if (!row.Opcode.EqualsIgnoreCase(OpCode.UseLimit) &&
                    !row.Opcode.EqualsIgnoreCase(OpCode.Characterize))
                {
                    break;
                }

                flowRows.Add(row.Copy());
            }
            return flowRows;
        }

        public void RemoveLimitRows(int index)
        {
            for (int i = index + 1; i < Rows.Count; i++)
            {
                FlowRow row = Rows[i];
                if (row.Opcode.EqualsIgnoreCase(OpCode.UseLimit) ||
                    row.Opcode.EqualsIgnoreCase(OpCode.Characterize))
                {
                    Rows.RemoveAt(i);
                    i--;
                }
                else
                {
                    break;
                }
            }
        }

        public bool IsSame(SubFlowSheet subFlowSheet)
        {
            if (Rows.Count != subFlowSheet.Rows.Count)
            {
                return false;
            }

            Type type = typeof(FlowRow);
            const BindingFlags memberFlags = BindingFlags.Public | BindingFlags.Instance;
            var members = new List<MemberInfo>();
            members.AddRange(type.GetFields(memberFlags).Where(p => p.DeclaringType == typeof(FlowRow)).ToList());

            for (int index = 0; index < Rows.Count; index++)
            {
                FlowRow sourceRow = Rows[index];
                if (sourceRow.Opcode.EqualsIgnoreCase(OpCode.Print))
                {
                    continue;
                }

                if (sourceRow.Parameter.ContainsIgnoreCase("_HEADER"))
                {
                    continue;
                }

                if (sourceRow.Parameter.ContainsIgnoreCase("_FOOTER"))
                {
                    continue;
                }

                FlowRow targetRow = subFlowSheet.Rows[index];
                foreach (MemberInfo t in members)
                {
                    var info = t as FieldInfo;
                    if (info != null)
                    {
                        FieldInfo fieldInfo = info;
                        if (fieldInfo.GetValue(sourceRow)!.ToString() != fieldInfo.GetValue(targetRow)!.ToString())
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        public void InsertBinTableBeforeReturnRow(List<BinTableRow> binTableRows)
        {
            var flowRows = new List<FlowRow>();
            var names = binTableRows.Select(x => x.Name).Distinct().ToList();
            foreach (string name in names)
            {
                var flowRow = new FlowRow { Opcode = OpCode.BinTable, Parameter = name };
                flowRows.Add(flowRow);
            }
            InsertBeforeReturnRow(flowRows);
        }

        public void InsertBeforeReturnRow(List<FlowRow> flowRows)
        {
            for (int index = Rows.Count - 1; index >= 0; index--)
            {
                FlowRow row = Rows[index];
                if (row.Opcode.EqualsIgnoreCase(OpCode.Return) && !row.IsBackup)
                {
                    Rows.InsertRange(index, flowRows);
                    return;
                }
            }
            Rows.AddRange(flowRows);
        }

        public void AddTestAndBinTableRow(string testItem, string job, string condition, string condName, string flag, bool bypassNeedBinTable, bool isBankRead = false, string enable = "", bool nopByEnableWord = false, string binTableEnableWd = "", string binOutStage = "")
        {
            var flowRow = new FlowRow
            {
                Enable = enable,
                Opcode = nopByEnableWord ? OpCode.Nop : OpCode.Test,
                Parameter = testItem,
                FailAction = string.IsNullOrEmpty(flag) ? RemoveHlnAndMrName(Combination.CombineByUnderLine("F", testItem)) : flag
            };
            if (!string.IsNullOrEmpty(condition))
            {
                flowRow.DeviceCondition = condition;
            }

            if (!string.IsNullOrEmpty(condName))
            {
                flowRow.DeviceName = condName;
            }

            if (!string.IsNullOrEmpty(job))
            {
                flowRow.Job = job;
            }

            Rows.Add(flowRow);

            if (!bypassNeedBinTable)
            {
                var binTableRow = new FlowRow { Enable = binTableEnableWd, Opcode = OpCode.BinTable, Parameter = string.IsNullOrEmpty(flag) ? RemoveHlnAndMrName(Combination.CombineByUnderLine("Bin_EFUSE", testItem)) : EfusePrefixRegex().Replace(flag, "").Replace("F_", "Bin_EFUSE_") };
                if (isBankRead)
                {
                    binTableRow.Parameter = binTableRow.Parameter.Replace("EFUSE", "EFUSE_RT");
                }

                if (!string.IsNullOrEmpty(binOutStage))
                {
                    binTableRow.Job = binOutStage;
                }
                else
                {
                    if (!string.IsNullOrEmpty(job))
                    {
                        binTableRow.Job = job;
                    }
                }

                Rows.Add(binTableRow);
            }
        }

        public void AddNopTestAndBinTableRow(string testItem, string job, string condition, string condName, string flag, bool needBinTable = true)
        {
            var flowRow = new FlowRow { Opcode = OpCode.Nop, Parameter = testItem, FailAction = string.IsNullOrEmpty(flag) ? RemoveHlnAndMrName(Combination.CombineByUnderLine("F", testItem)) : flag };
            if (!string.IsNullOrEmpty(condition))
            {
                flowRow.DeviceCondition = condition;
            }

            if (!string.IsNullOrEmpty(condName))
            {
                flowRow.DeviceName = condName;
            }

            if (!string.IsNullOrEmpty(job))
            {
                flowRow.Job = job;
            }

            Rows.Add(flowRow);

            if (needBinTable)
            {
                var binTableRow = new FlowRow { Opcode = OpCode.BinTable, Parameter = string.IsNullOrEmpty(flag) ? RemoveHlnAndMrName(Combination.CombineByUnderLine("Bin_EFUSE", testItem)) : flag.Replace("F_", "Bin_EFUSE_") };
                if (!string.IsNullOrEmpty(job))
                {
                    binTableRow.Job = job;
                }

                Rows.Add(binTableRow);
            }
        }

        protected StringBuilder? WriteSheet(string version, SheetInfo sheetInfo)
        {
            if (Rows.Count == 0)
            {
                return null;
            }

            var sb = new StringBuilder();
            //var sw = new StreamWriter(fileName, false);
            {
                int maxCount = GetMaxCount(sheetInfo);
                int labelIndex = GetIndexFrom(sheetInfo, "Label");
                int enableIndex = GetIndexFrom(sheetInfo, "Enable");
                int jobIndex = GetIndexFrom(sheetInfo, "Gate", "Job");
                int partIndex = GetIndexFrom(sheetInfo, "Gate", "Part");
                int envIndex = GetIndexFrom(sheetInfo, "Gate", "Env");
                int opcodeIndex = GetIndexFrom(sheetInfo, "Command", "Opcode");
                int parameterIndex = GetIndexFrom(sheetInfo, "Command", "Parameter");
                int tNameIndex = GetIndexFrom(sheetInfo, "TName");
                int tNumIndex = GetIndexFrom(sheetInfo, "TNum");
                int loLimIndex = GetIndexFrom(sheetInfo, "Limits", "LoLim");
                int hiLimIndex = GetIndexFrom(sheetInfo, "Limits", "HiLim");
                int scaleIndex = GetIndexFrom(sheetInfo, "Datalog Display Results", "Scale");
                int unitsIndex = GetIndexFrom(sheetInfo, "Datalog Display Results", "Units");
                int formatIndex = GetIndexFrom(sheetInfo, "Datalog Display Results", "Format");
                int binNumberPassIndex = GetIndexFrom(sheetInfo, "Bin Number", "Pass");
                int binNumberFailIndex = GetIndexFrom(sheetInfo, "Bin Number", "Fail");
                int sortNumberPassIndex = GetIndexFrom(sheetInfo, "Sort Number", "Pass");
                int sortNumberFailIndex = GetIndexFrom(sheetInfo, "Sort Number", "Fail");
                int resultIndex = GetIndexFrom(sheetInfo, "Result");
                int actionPassIndex = GetIndexFrom(sheetInfo, "Action", "Pass");
                int actionFailIndex = GetIndexFrom(sheetInfo, "Action", "Fail");
                int stateIndex = GetIndexFrom(sheetInfo, "State");
                int groupSpecifierIndex = GetIndexFrom(sheetInfo, "Group", "Specifier");
                int groupSenseIndex = GetIndexFrom(sheetInfo, "Group", "Sense");
                int groupConditionIndex = GetIndexFrom(sheetInfo, "Group", "Condition");
                int groupNameIndex = GetIndexFrom(sheetInfo, "Group", "Name");
                int deviceSenseIndex = GetIndexFrom(sheetInfo, "Device", "Sense");
                int deviceConditionIndex = GetIndexFrom(sheetInfo, "Device", "Condition");
                int deviceNameIndex = GetIndexFrom(sheetInfo, "Device", "Name");
                int debugAssumeIndex = GetIndexFrom(sheetInfo, "Debug", "Assume");
                int debugSitesIndex = GetIndexFrom(sheetInfo, "Debug", "Sites");
                int elapsedTimeIndex = GetIndexFrom(sheetInfo, "CT Profile Data", "Elapsed Time (s)");
                int backgroundTypeIndex = GetIndexFrom(sheetInfo, "CT Profile Data", "Background Type");
                int serializeIndex = GetIndexFrom(sheetInfo, "CT Profile Data", "Serialize");
                int resourceLockIndex = GetIndexFrom(sheetInfo, "CT Profile Data", "Resource Lock");
                int flowStepLockedIndex = GetIndexFrom(sheetInfo, "CT Profile Data", "Flow Step Locked");
                int commentIndex = GetIndexFrom(sheetInfo, "Comment");

                #region headers
                int startRow = sheetInfo.Columns.RowCount;
                sb.AppendLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                for (int i = 1; i < startRow; i++)
                {
                    string[] arr = [.. Enumerable.Repeat("", maxCount)];

                    SetField(sheetInfo, i, arr);

                    SetColumns(sheetInfo, i, arr);

                    WriteHeader(arr, sb);
                }
                #endregion

                #region data
                var flowRows = Rows.Where(x => !x.IsBackup).ToList();
                WriteFlowRows(flowRows, maxCount, labelIndex, enableIndex, jobIndex, partIndex, envIndex, opcodeIndex, parameterIndex, tNameIndex, tNumIndex, loLimIndex, hiLimIndex, scaleIndex, unitsIndex, formatIndex, binNumberPassIndex, binNumberFailIndex, sortNumberPassIndex, sortNumberFailIndex, resultIndex, actionPassIndex, actionFailIndex, stateIndex, groupSpecifierIndex, groupSenseIndex, groupConditionIndex, groupNameIndex, deviceSenseIndex, deviceConditionIndex, deviceNameIndex, debugAssumeIndex, debugSitesIndex, elapsedTimeIndex, backgroundTypeIndex, serializeIndex, resourceLockIndex, flowStepLockedIndex, commentIndex, sb);

                var backupRows = Rows.Where(x => x.IsBackup).ToList();
                if (backupRows.Count != 0)
                {
                    PrintBackupEmptyRows(sb);
                    WriteFlowRows(backupRows, maxCount, labelIndex, enableIndex, jobIndex, partIndex, envIndex, opcodeIndex, parameterIndex, tNameIndex, tNumIndex, loLimIndex, hiLimIndex, scaleIndex, unitsIndex, formatIndex, binNumberPassIndex, binNumberFailIndex, sortNumberPassIndex, sortNumberFailIndex, resultIndex, actionPassIndex, actionFailIndex, stateIndex, groupSpecifierIndex, groupSenseIndex, groupConditionIndex, groupNameIndex, deviceSenseIndex, deviceConditionIndex, deviceNameIndex, debugAssumeIndex, debugSitesIndex, elapsedTimeIndex, backgroundTypeIndex, serializeIndex, resourceLockIndex, flowStepLockedIndex, commentIndex, sb);
                }
                #endregion
            }
            return sb;
        }

        private static void WriteFlowRows(List<FlowRow> flowRows, int maxCount, int labelIndex, int enableIndex, int jobIndex, int partIndex, int envIndex, int opcodeIndex, int parameterIndex, int tNameIndex, int tNumIndex, int loLimIndex, int hiLimIndex, int scaleIndex, int unitsIndex, int formatIndex, int binNumberPassIndex, int binNumberFailIndex, int sortNumberPassIndex, int sortNumberFailIndex, int resultIndex, int actionPassIndex, int actionFailIndex, int stateIndex, int groupSpecifierIndex, int groupSenseIndex, int groupConditionIndex, int groupNameIndex, int deviceSenseIndex, int deviceConditionIndex, int deviceNameIndex, int debugAssumeIndex, int debugSitesIndex, int elapsedTimeIndex, int backgroundTypeIndex, int serializeIndex, int resourceLockIndex, int flowStepLockedIndex, int commentIndex, StringBuilder stringBuilder)
        {
            for (int index = 0; index < flowRows.Count; index++)
            {
                FlowRow row = flowRows[index];
                string[] arr = [.. Enumerable.Repeat("", maxCount)];
                arr[0] = row.ColumnA ?? "";
                arr[labelIndex] = row.Label;
                arr[enableIndex] = row.Enable;
                arr[jobIndex] = row.Job;
                arr[partIndex] = row.Part;
                arr[envIndex] = row.Env;
                arr[opcodeIndex] = row.Opcode;
                arr[parameterIndex] = row.Parameter;
                arr[tNameIndex] = row.TName;
                arr[tNumIndex] = row.TNum;
                arr[loLimIndex] = row.LoLim;
                arr[hiLimIndex] = row.HiLim;
                arr[scaleIndex] = row.Scale;
                arr[unitsIndex] = row.Units;
                arr[formatIndex] = row.Format;
                arr[binNumberPassIndex] = row.BinPass;
                arr[binNumberFailIndex] = row.BinFail;
                arr[sortNumberPassIndex] = row.SortPass;
                arr[sortNumberFailIndex] = row.SortFail;
                arr[resultIndex] = row.Result;
                arr[actionPassIndex] = row.PassAction;
                arr[actionFailIndex] = row.FailAction;
                arr[stateIndex] = row.State;
                arr[groupSpecifierIndex] = row.GroupSpecifier;
                arr[groupSenseIndex] = row.GroupSense;
                arr[groupConditionIndex] = row.GroupCondition;
                arr[groupNameIndex] = row.GroupName;
                arr[deviceSenseIndex] = row.DeviceSense;
                arr[deviceConditionIndex] = row.DeviceCondition;
                arr[deviceNameIndex] = row.DeviceName;
                arr[debugAssumeIndex] = row.DebugAssume;
                arr[debugSitesIndex] = row.DebugSites;
                arr[elapsedTimeIndex] = row.CtProfileDataElapsedTime;
                arr[backgroundTypeIndex] = row.CtProfileDataBackgroundType;
                arr[serializeIndex] = row.CtProfileDataSerialize;
                arr[resourceLockIndex] = row.CtProfileDataResourceLock;
                arr[flowStepLockedIndex] = row.CtProfileDataFlowStepLocked;
                arr[commentIndex] = row.Comment;

                string str = row.Comment1.Trim('\t');
                row.Comment1 = string.IsNullOrEmpty(str) ? str : row.Comment1.TrimEnd('\t');

                if (string.IsNullOrEmpty(row.Comment1))
                {
                    stringBuilder.AppendLine(string.Join("\t", arr) + "\t");
                }
                else
                {
                    stringBuilder.AppendLine(string.Join("\t", arr) + "\t" + row.Comment1 + "\t");
                }
            }
        }

        public List<string> GetUsedFlowSheets(List<SubFlowSheet> subFlowSheets)
        {
            var sheets = new List<string>();
            foreach (FlowRow flowRow in Rows)
            {
                if (flowRow == null)
                {
                    continue;
                }

                if (flowRow.Opcode.EqualsIgnoreCase(OpCode.Call))
                {
                    if (subFlowSheets.Exists(x => x.Name.EqualsIgnoreCase(flowRow.Parameter)))
                    {
                        sheets.Add(flowRow.Parameter);
                        SubFlowSheet sheet = subFlowSheets.Find(x => x.Name.EqualsIgnoreCase(flowRow.Parameter))!;
                        sheets.AddRange(sheet.GetUsedFlowSheets(subFlowSheets));
                    }
                }
            }
            return [.. sheets.Distinct()];
        }

        public bool IsMatchFlowRow(FlowRow flowRow)
        {
            foreach (FlowRow row in Rows)
            {
                if (row.Opcode.EqualsIgnoreCase(flowRow.Opcode) &&
                    row.Parameter.EqualsIgnoreCase(flowRow.Parameter) &&
                    row.Enable.EqualsIgnoreCase(flowRow.Enable) &&
                    !row.IsBackup)
                {
                    return true;
                }
            }
            return false;
        }

        public FlowRow? ReplaceFlowRow(FlowRow flowRow)
        {
            for (int index = 0; index < Rows.Count; index++)
            {
                FlowRow row = Rows[index];
                if (row.Opcode.EqualsIgnoreCase(flowRow.Opcode) &&
                    row.Parameter.EqualsIgnoreCase(flowRow.Parameter) &&
                    row.Enable.EqualsIgnoreCase(flowRow.Enable) &&
                    !row.IsBackup)
                {
                    InsertRow(index, flowRow);
                    return row;
                }
            }
            return null;
        }

        #region RemoveTheSameIf
        public void RemoveTheSameIf()
        {
            var condition = new IfCondition();
            var removeList = new List<int>();
            for (int index = 0; index < Rows.Count; index++)
            {
                FlowRow flowRow = Rows[index];
                if (flowRow.Opcode.EqualsIgnoreCase(OpCode.If))
                {
                    IfCondition ifCondition = GetIfCondition([.. Rows], index);
                    if (condition.EndIfIndex == index - 1 && condition.If.EqualsIgnoreCase(ifCondition.If))
                    {
                        if (ifCondition.ElseIfIndex == -1 &&
                            ifCondition.ElseIndex == -1 &&
                            condition.ElseIfIndex == -1 &&
                            condition.ElseIndex == -1)
                        {
                            removeList.Add(condition.EndIfIndex);
                            removeList.Add(index);
                        }
                    }
                    condition = ifCondition;
                }
            }

            if (removeList.Count != 0)
            {
                for (int index = removeList.Count - 1; index >= 0; index--)
                {
                    int remove = removeList[index];
                    Rows.RemoveAt(remove);
                }
            }
        }

        private IfCondition GetIfCondition(FlowRows flowRows, int index)
        {
            List<string> parameterParts = [Rows[index].Parameter, Rows[index].Enable, Rows[index].Job];
            string parameterKey = string.Join("_", parameterParts);

            var ifCondition = new IfCondition { If = parameterKey, IfIndex = index };
            for (int i = index; i < flowRows.Count; i++)
            {
                FlowRow flowRow = Rows[i];
                if (flowRow.Opcode.EqualsIgnoreCase(OpCode.If))
                {
                    if (i != index)
                    {
                        i = SkipRow(flowRows, i + 1);
                    }

                    if (i == -1)
                    {
                        return ifCondition;
                    }
                }
                else if (flowRow.Opcode.EqualsIgnoreCase(OpCode.ElseIf))
                {
                    ifCondition.ElseIf = parameterKey;
                    ifCondition.ElseIfIndex = i;
                }
                else if (flowRow.Opcode.EqualsIgnoreCase(OpCode.Else))
                {
                    ifCondition.Else = parameterKey;
                    ifCondition.ElseIndex = i;
                }
                else if (flowRow.Opcode.EqualsIgnoreCase(OpCode.EndIf))
                {
                    ifCondition.EndIf = parameterKey;
                    ifCondition.EndIfIndex = i;
                    break;
                }
            }
            return ifCondition;
        }

        private int SkipRow(FlowRows flowRows, int index)
        {
            int cnt = 0;
            for (int i = index; i < flowRows.Count; i++)
            {
                FlowRow flowRow = Rows[i];
                if (flowRow.Opcode.EqualsIgnoreCase(OpCode.If))
                {
                    cnt++;
                }
                else if (flowRow.Opcode.EqualsIgnoreCase(OpCode.EndIf))
                {
                    if (cnt == 0)
                    {
                        return i;
                    }

                    cnt--;
                }
            }
            return -1;
        }
        #endregion

        private static string RemoveHlnAndMrName(string testName)
        {
            if (string.IsNullOrEmpty(testName))
            {
                return testName;
            }

            return testName.Replace("_MR0", "").Replace("_MR1", "").Replace("_HV", "").Replace("_LV", "").Replace("_NV", "");
        }

        public void FilterFlowJobs(List<string> allJobsHardIp)
        {
            foreach (FlowRow row in Rows)
            {
                if (!string.IsNullOrEmpty(row.Job))
                {
                    List<string> jobs = [.. row.Job.ToUpper().Split(',')];
                    if (jobs.All(allJobsHardIp.Contains) && allJobsHardIp.All(jobs.Contains))
                    {
                        row.Job = "";
                    }
                }
            }
        }
    }
}
