using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.DividerManager;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public class IdsBlockInsGenerator : BlockInstanceGenerator
    {
        public IdsBlockInsGenerator(HardIpInputData hardIpInputData, string sheetName, HardIpSheet hardIpSheet)
            : base(hardIpInputData, sheetName, hardIpSheet)
        {
            InstanceRowGenerator = new IdsInsRowGenerator(hardIpInputData, hardIpSheet, sheetName);
        }

        /// <summary>
        /// Generate instance rows for IDS sheet and insert these rows to "TestInst_DCTEST_IDS"
        /// </summary>
        public override List<InstanceSheet> GenBlockInsRows()
        {
            List<InstanceSheet> instanceSheetList = new List<InstanceSheet>();
            var charInsSheet = new InstanceSheet(HardIpConstData.PrefixInsSheetByVoltage + HardIpConstData.LabelChar);
            var ids = new InstanceSheet("TestInst_" + SheetName, SheetName);
            List<HardIpPattern> patLstToGen = DividerMain.DivideInstancePattern(HardIpInputData, HardIpSheet.Rows).ToList();
            foreach (HardIpPattern pattern in patLstToGen)
            {
                InstanceRowGenerator.LabelVoltage = "";
                InstanceRowGenerator.Pat = pattern;
                List<InstanceRow> insRowList = InstanceRowGenerator.GenInsRows();
                if (pattern.SkipDotNet)
                {
                    continue;
                }

                foreach (InstanceRow insRow in insRowList)
                {
                    if (InstanceRowGenerator.Pat.ForceCondition.IsShmooInCharInst)
                    {
                        charInsSheet.AddRow(insRow);
                    }
                    else
                    {
                        ids.AddRow(insRow);
                    }
                }
            }
            if (charInsSheet.Rows.Count != 0)
            {
                instanceSheetList.Add(charInsSheet);
            }

            if (ids.Rows.Count != 0)
            {
                instanceSheetList.Add(ids);
            }

            return instanceSheetList;
        }
    }
}
