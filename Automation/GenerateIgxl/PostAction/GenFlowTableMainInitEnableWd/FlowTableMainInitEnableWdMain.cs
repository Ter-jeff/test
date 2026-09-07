using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Static;

using IgxlLib;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.PostAction.GenFlowTableMainInitEnableWd
{
    public class FlowTableMainInitEnableWdMain
    {
        public void WorkFlow()
        {
            var flags = TestProgram.IgxlWorkBk.BinTblSheets.SelectMany(x => x.Value.Rows)
                .SelectMany(x => x.ItemList.Split(',')).Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            var existFlags = TestProgram.IgxlWorkBk.SubFlowSheets.SelectMany(x => x.Value.Rows)
                .SelectMany(x => x.FailAction.Split(',')).Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            var extraflags = flags.Except(existFlags, StringComparer.CurrentCultureIgnoreCase).ToList();
            flags.AddRange(extraflags);
            if (TestProgram.IgxlWorkBk.GetInstanceSheet("TestInst_Non_Bincut") != null)
            {
                var pinGrps = TestProgram.IgxlWorkBk.GetInstanceSheet("TestInst_Non_Bincut").Rows.Where(x => x.GetPinGrpFlags() != null).Select(x => x.GetPinGrpFlags()).ToList();
                var pinGrpFlags = new List<string>();
                foreach (string flag in pinGrps.SelectMany(pinGrp => pinGrp.Where(flag => !pinGrpFlags.Exists(x => x.Equals(flag, StringComparison.OrdinalIgnoreCase)))))
                {
                    pinGrpFlags.Add(flag);
                }
                TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(pinGrpFlags, "Harvest", FolderStructure.DirMain);
            }
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(flags, "Autogen", FolderStructure.DirMain);


            ExcelWorkbook workbook = SettingStatic.BasicConfigWorkbook;
            SubFlowSheet oldSheet = null;
            foreach (ExcelWorksheet workSheet in workbook.Worksheets)
            {
                if (workSheet.Name.Equals(IgxlWorkBook.FlowTableMainInitEnableWd, StringComparison.CurrentCultureIgnoreCase))
                {
                    var flowReader = new ReadFlowSheet();
                    oldSheet = flowReader.ReadSheet(workSheet);
                }
            }
            if (oldSheet != null)
            {
                SubFlowSheet mainInitEnableWdSheet = TestProgram.IgxlWorkBk.GetFlowSheet(IgxlWorkBook.FlowTableMainInitEnableWd, FolderStructure.DirMain);
                if (mainInitEnableWdSheet != null)
                {
                    foreach (FlowRow row in mainInitEnableWdSheet.Rows)
                    {
                        if (oldSheet.IsMatchFlowRow(row))
                        {
                            FlowRow oldFlowRow = oldSheet.ReplaceFlowRow(row);
                            oldFlowRow.IsBackup = true;
                        }
                        else
                        {
                            InsertBeforeReturnRow(oldSheet, row);
                        }
                    }
                    AddCommentFlags(oldSheet);
                    GenerateScanEnableWdFlow(oldSheet);
                    TestProgram.IgxlWorkBk.SetFlowSheet(IgxlWorkBook.FlowTableMainInitEnableWd, oldSheet);
                }
            }
        }
        private void AddCommentFlags(SubFlowSheet oldSheet)
        {
            var insertRows = new List<FlowRow>();
            var flagsTrueList = new List<string>
            {
                "F_PrintHarvReport"
            };
            var flagsFalseList = new List<string>
            {
                "F_Debug_all"
            };
            var enableWords = new List<string>
            {
                "DebugPrint","Vdd_Binning_PTE_Debug","Debug_BinCutCOF_Stored","Enable_SFC","Enable_RSCR","Enable_FingerPrint","Enable_FFC","ForceUpdateJob", "LogLevel_Debug", "LogLevel_NPI", "LogLevel_MP", "CurrentProfile_Enhance","DoNotCheckBuildConfig","DebugPrint_Detail"
            };
            if (flagsFalseList.Any())
            {
                insertRows.Add(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = string.Join(",", flagsFalseList), ColumnA = "Autogen Default" });
            }

            if (flagsTrueList.Any())
            {
                insertRows.Add(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = string.Join(",", flagsTrueList), ColumnA = "Autogen Default" });
            }

            foreach (string enableword in enableWords)
            {
                insertRows.Add(new FlowRow { Opcode = OpCode.Nop, Enable = enableword, ColumnA = "Autogen Default" });
            }
            foreach (FlowRow row in insertRows)
            {
                InsertBeforeReturnRow(oldSheet, row);
            }
        }
        private void InsertBeforeReturnRow(SubFlowSheet subFlowSheet, FlowRow flowRow)
        {
            int index = subFlowSheet.Rows.FindLastIndex(x => x.Opcode.Equals("print", StringComparison.OrdinalIgnoreCase) &&
                                                             Regex.IsMatch(x.Parameter, @"\w*Flow_Table_Main_Init_EnableWd Stop\w*"));
            if (index != -1)
            {
                subFlowSheet.InsertRow(index, flowRow);
            }
            else
            {
                subFlowSheet.InsertBeforeReturnRow(new List<FlowRow> { flowRow });
            }
        }

        private void GenerateScanEnableWdFlow(SubFlowSheet oldSheet)
        {
            if (TestPlanStatic.ScanEnableWordSheet != null && oldSheet != null)
            {
                string sheetName = TestPlanStatic.ScanEnableWordSheet.SheetName.Equals("Scan_Enable_words", StringComparison.OrdinalIgnoreCase) ? "Flow_ScanEnableWdSheet" : "Flow_EnableWdSheet";
                var scanEnableSubFlow = new SubFlowSheet(sheetName);
                string headerFooterName = Regex.Replace(sheetName, "^Flow_", "", RegexOptions.IgnoreCase);
                scanEnableSubFlow.AddPrintStartRow(headerFooterName);
                var flowRow = new FlowRow { Opcode = OpCode.Call, Parameter = sheetName };
                Dictionary<string, string> flagNameDic = TestPlanStatic.ScanEnableWordSheet.FlagNameDic;
                foreach (KeyValuePair<string, Dictionary<string, string>> enablewd in TestPlanStatic.ScanEnableWordSheet.EnableWordsDic)
                {
                    scanEnableSubFlow.AddRow(new FlowRow { Enable = enablewd.Key, Opcode = OpCode.Nop });
                    var useJobList = enablewd.Value.Where(x => x.Value.Equals("V", StringComparison.OrdinalIgnoreCase)).Select(x => x.Key).ToList();
                    if (useJobList.Any())
                    {
                        scanEnableSubFlow.AddRow(new FlowRow { Opcode = OpCode.EnableFlowWd, Parameter = enablewd.Key, Job = string.Join(",", useJobList) });
                    }
                    var unuseJobList = enablewd.Value.Where(x => x.Value.Equals("X", StringComparison.OrdinalIgnoreCase)).Select(x => x.Key).ToList();
                    if (unuseJobList.Any())
                    {
                        scanEnableSubFlow.AddRow(new FlowRow { Opcode = OpCode.DisableFlowWd, Parameter = enablewd.Key, Job = string.Join(",", unuseJobList) });
                    }
                    if (flagNameDic.TryGetValue(enablewd.Key, out string flagName))
                    {
                        scanEnableSubFlow.AddRow(new FlowRow { Opcode = OpCode.FlagFalse, Parameter = flagName });
                        scanEnableSubFlow.AddRow(new FlowRow { Opcode = OpCode.FlagTrue, Parameter = flagName, Enable = enablewd.Key });
                    }
                }
                scanEnableSubFlow.AddPrintEndRow(headerFooterName);
                scanEnableSubFlow.AddReturnRow();
                InsertBeforeReturnRow(oldSheet, flowRow);
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirMain, scanEnableSubFlow);
            }
        }
    }
}
