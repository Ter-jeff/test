using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.DividerManager;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class RtosBlockInsGenerator : BlockInstanceGenerator
    {
        public RtosBlockInsGenerator(HardIpInputData hardIpInputData, string sheetName, HardIpSheet hardIpSheet) : base(hardIpInputData, sheetName, hardIpSheet)
        {
            InstanceRowGenerator = new RtosInsRowGenerator(hardIpInputData, hardIpSheet, sheetName);
        }

        /// <summary>
        /// Generate instance rows for IDS sheet and insert these rows to "TestInst_DCTEST_IDS"
        /// </summary>
        public override List<InstanceSheet> GenBlockInsRows()
        {
            List<InstanceSheet> instanceSheets = new List<InstanceSheet>();
            var instanceSheet = new InstanceSheet("TestInst_" + SheetName, SheetName);
            List<HardIpPattern> hardIpPatterns = DividerMain.DivideInstancePattern(HardIpInputData, HardIpSheet.Rows).ToList();
            foreach (HardIpPattern hardIpPattern in hardIpPatterns)
            {
                InstanceRowGenerator.LabelVoltage = "";
                InstanceRowGenerator.Pat = hardIpPattern;
                List<InstanceRow> instanceRows = InstanceRowGenerator.GenInsRows();
                foreach (InstanceRow instanceRow in instanceRows)
                {
                    instanceSheet.AddRow(instanceRow);
                }
            }

            if (instanceSheet.Rows.Count != 0)
            {
                instanceSheets.Add(instanceSheet);
            }

            return instanceSheets;
        }
    }
}
