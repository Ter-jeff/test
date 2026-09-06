using System;
using System.IO;
using System.Linq;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using CommonLib.Extension;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class GenMainFlow
    {
        // added by Jackie on 2016/11/30, to add Flow_Char into MainFlow
        public static void WorkFlow()
        {
            string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.MainFolder);

            GeneralFunc.WriteMessage("Update main flow sheet... ");

            JobListSheet jobSheet = LocalSpecs.TestProgram.JoblistSheet;

            foreach (JobRow job in jobSheet.Rows)
            {
                SubFlowSheet mainFlow = LocalSpecs.TestProgram.GetMainFlowSheet(job);

                FlowRow charPlanFromTP = mainFlow.Rows.FirstOrDefault(p => !string.IsNullOrEmpty(p.ColumnA) &&
                    Path.GetFileName(LocalSpecs.InputParam.CharPlan).EqualsIgnoreCase(p.ColumnA));

                int index = mainFlow.Rows.FindIndex(x => x.Parameter
                    .Equals("Flow_DCTEST_IDS", StringComparison.OrdinalIgnoreCase));
                if (index == -1)
                {
                    index = mainFlow.Rows.FindIndex(x => x.Parameter
                        .Equals("Flow_DC_Conti", StringComparison.OrdinalIgnoreCase));
                }

                var row = new FlowRow { Opcode = "call", Parameter = "Flow_Char" };

                if (LocalSpecs.FileStructure.Count == 0)    // DFTL flow
                {
                    if (charPlanFromTP != null)
                    {
                        if (charPlanFromTP.Opcode == "nop")
                        {
                            charPlanFromTP.Opcode = "call";
                        }
                    }
                }
                else    // Individual flow
                {
                    if (index == -1)
                    {
                        mainFlow.Rows.Add(row);
                    }
                    else
                    {
                        mainFlow.Rows.Insert(index + 1, row);

                        int firstBackupRow = -1;
                        for (int i = index + 2; i < mainFlow.Rows.Count; i++)
                        {
                            if (mainFlow.Rows[i].Opcode != "set-device")
                            {
                                FlowRow flowRow = mainFlow.Rows[i];
                                flowRow.Opcode = "nop";
                            }

                            if (firstBackupRow == -1 && mainFlow.Rows[i].IsBackup)
                            {
                                firstBackupRow = i;
                            }
                        }
                        if (firstBackupRow != -1)
                        {
                            mainFlow.Rows.Insert(firstBackupRow, new FlowRow());
                        }
                    }
                }

                string outputPath = Path.Combine(outputFolder, mainFlow.Name + ".txt");
                mainFlow.Write(outputPath, LocalSpecs.ExportVersion < 9.0 ? "2.3" : "3.0");
                LocalSpecs.GenSheets.Add(mainFlow);
            }
        }
    }
}
