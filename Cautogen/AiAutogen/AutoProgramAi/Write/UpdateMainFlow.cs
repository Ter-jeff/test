using System;
using System.Text.RegularExpressions;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateMainFlow
    {
        public SubFlowSheet Work(SubFlowSheet mainFlow, string flowSheet)
        {
            //var index = mainFlow.FlowRows.FindLastIndex(x => x.Parameter.Equals("Flow_DCTEST_IDS", StringComparison.OrdinalIgnoreCase) && x.IsBackup == false);
            var index = mainFlow.Rows.FindLastIndex(x => Regex.IsMatch(x.Parameter, "Flow_DCTEST_IDS*.*") && x.IsBackup == false);
            if (index == -1)
                index = mainFlow.Rows.FindLastIndex(x => x.Opcode
                    .Equals("set-device", StringComparison.OrdinalIgnoreCase) && x.IsBackup == false) - 1;

            var row = new FlowRow { Opcode = "call", Parameter = flowSheet };
            if (index < 0)
            {
                mainFlow.Rows.Add(row);
            }
            else
            {
                mainFlow.Rows.Insert(index + 1, row);
                var firstBackupRow = -1;
                for (var i = index + 2; i < mainFlow.Rows.Count; i++)
                {
                    if (mainFlow.Rows[i].Opcode != "set-device")
                    {
                        var flowRow = mainFlow.Rows[i];
                        flowRow.Opcode = "nop";
                    }

                    if (firstBackupRow == -1 && mainFlow.Rows[i].IsBackup)
                        firstBackupRow = i;
                }
                if (firstBackupRow != -1)
                    mainFlow.Rows.Insert(firstBackupRow, new FlowRow());
            }

            return mainFlow;
        }
    }
}
