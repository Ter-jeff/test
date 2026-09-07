using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Concurrent;

namespace Automation.GenerateIgxl.PostAction.ConcurrentSequence
{
    public class ConcurrentSequenceMain
    {
        public void WorkFlow(ConcurrentFlowSheet concurrentFlow, Dictionary<string, SubFlowSheet> subFlowSheets, Dictionary<string, InstanceSheet> insSheets)
        {
            try
            {
                Response.Report("Generating Concurrent Sequence sheet ...", percentage: 5);
                var concurrentSequenceSheet = new ConcurrentSequenceSheet("Concurrent Sequence");
                Dictionary<string, bool> instDictionary = ParseDictionary(insSheets);
                foreach (ConcurrentFlowSheetRow sequence in concurrentFlow.Rows)
                {
                    var concurrentRow = new ConcurrentSequenceRow { SequenceName = sequence.SequenceName };
                    foreach (string subflow in sequence.Subflows)
                    {
                        string path = subFlowSheets.Keys.FirstOrDefault(x => x.ContainsIgnoreCase(subflow.ToUpper()));
                        if (path == null)
                        {
                            continue;
                        }

                        bool result = subFlowSheets.TryGetValue(path, out SubFlowSheet subflowSheet);
                        if (result)
                        {
                            var flowStep = new List<FlowStep>();
                            for (int i = 0; i < subflowSheet.Rows.Count; i++)
                            {
                                bool hasPattern = IsPatternTest(subflowSheet.Rows[i], instDictionary);
                                if (subflowSheet.Rows[i].Opcode.Equals(OpCode.Test))
                                {
                                    flowStep.Add(new FlowStep { FlowStepItem = "=" + subflow + "!H" + (i + 5), BackgroundSubStep = hasPattern ? "Pattern" : "None" });
                                }
                            }
                            concurrentRow.FlowSteps.Add(subflow, flowStep);
                        }
                    }
                    concurrentSequenceSheet.AddRow(concurrentRow);
                }
                Write(concurrentSequenceSheet);
                Response.Report("Concurrent Sequence Completed!", percentage: 100);
            }
            catch (Exception e)
            {
                Response.Report("Concurrent AutoGen Failed: " + e.StackTrace, EnumMessageLevel.Error, 100);
            }
        }

        private Dictionary<string, bool> ParseDictionary(Dictionary<string, InstanceSheet> insSheets)
        {
            var dict = new Dictionary<string, bool>();
            foreach (KeyValuePair<string, InstanceSheet> inst in insSheets)
            {
                string instSheetName = Path.GetFileName(inst.Key);
                if (instSheetName.StartsWith("Inst_") || instSheetName.StartsWith("TestInst_"))
                {
                    foreach (InstanceRow row in inst.Value.Rows)
                    {
                        if (!dict.ContainsKey(row.TestName))
                        {
                            dict.Add(row.TestName, !string.IsNullOrEmpty(row.GetArgument("patterns")) || !string.IsNullOrEmpty(row.GetArgument("patset")) || !string.IsNullOrEmpty(row.GetArgument("pat")));
                        }
                    }
                }
            }

            return dict;
        }

        private bool IsPatternTest(FlowRow flowRow, Dictionary<string, bool> isPatDictionary)
        {
            if (isPatDictionary.TryGetValue(flowRow.Parameter, out bool test))
            {
                return test;
            }

            return !(flowRow.Parameter.ContainsIgnoreCase("HEADER") || flowRow.Parameter.ContainsIgnoreCase("FOOTER") ||
                     flowRow.Parameter.ToUpper().StartsWith("RELAY"));
        }

        private void Write(ConcurrentSequenceSheet concurrentSequence)
        {
            var excel = new ExcelPackage(new FileInfo(Path.Combine(FolderStructure.DirCommon, "Concurrent Sequence.xlsx")));
            ExcelWorksheet workSheet = excel.Workbook.AddSheet("Concurrent Sequence");
            var headers = new Dictionary<string, int> { { "Disable", 2 }, { "Sequence Name", 3 }, { "Flow Steps", 4 } };
            int curIdx = 4;
            for (int i = 0; i < concurrentSequence.Rows.Max(x => x.FlowSteps.Count); i++)
            {
                headers.Add("FlowstepItem" + (i + 1), ++curIdx);
                headers.Add("BackgroundSubStep" + (i + 1), ++curIdx);
            }

            workSheet.Cells[1, 1].Value = "DTCTSequencesSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1";
            workSheet.Cells[1, 2].Value = "Concurrent Sequence";
            foreach (KeyValuePair<string, int> header in headers)
            {
                workSheet.Cells[3, header.Value].Value = header.Key;
            }
            int curRow = 4;
            foreach (ConcurrentSequenceRow row in concurrentSequence.Rows)
            {
                for (int i = 0; i < row.FlowSteps.Keys.Count; i++)
                {
                    workSheet.Cells[curRow, headers["Sequence Name"]].Value = row.SequenceName;
                    workSheet.Cells[curRow, headers["Flow Steps"]].Value = row.FlowSteps.Keys.ElementAt(i);
                    curRow++;
                }
                for (int i = 0; i < row.FlowSteps.Keys.Count; i++)
                {
                    for (int j = 0; j < row.FlowSteps.Values.ElementAt(i).Count; j++)
                    {
                        workSheet.Cells[curRow, headers["Sequence Name"]].Value = row.SequenceName;
                        workSheet.Cells[curRow, headers["FlowstepItem" + (i + 1)]].Value = row.FlowSteps.Values.ElementAt(i)[j].FlowStepItem;
                        workSheet.Cells[curRow, headers["BackgroundSubStep" + (i + 1)]].Value = row.FlowSteps.Values.ElementAt(i)[j].BackgroundSubStep;
                        curRow++;
                    }
                }
            }

            excel.ExportWorkBook2Txt(FolderStructure.DirCommon);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommon, "Concurrent Sequence");
        }
    }
}
