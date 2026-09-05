using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.Relay.RelayConst;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Static;

namespace Automation.GenerateIgxl.PostAction.Relay
{
    public class RelayMain
    {
        public List<RelayItemNew> WorkFlow()
        {
            Response.Report("Running Relay PreCheck ...", EnumMessageLevel.General, 10);

            var reader = new RelayReader();
            RelaySheet sheet = reader.ReadSheet(EpWorkbook.TestPlanWorkbook.Worksheets["Relay"]);

            Response.Report("Parsing Relay Table ...", EnumMessageLevel.General, 30);
            //AutoGen(para, nWireFlowSheets);
            string commonSheetName = TestProgram.IgxlWorkBk.InsSheets.Keys.ToList().FirstOrDefault(x => x.Contains(Path.DirectorySeparatorChar + "TestInst_Common"));
            if (commonSheetName != null)
            {
                InstanceSheet commonSheet = TestProgram.IgxlWorkBk.InsSheets[commonSheetName];
                if (commonSheet != null)
                {
                    foreach (RelayItemNew item in sheet.RelayItems)
                    {
                        InstanceRow instRow = new InstanceRow();
                        OutputConst relayFunction = new OutputConst(string.Join(",", item.On), string.Join(",", item.Off));
                        instRow.TestName = item.Module;
                        instRow.VbtType = relayFunction.Function.Type;
                        instRow.VbtName = relayFunction.Function.FullFunctionName;
                        instRow.ArgList = relayFunction.Function.Parameters;
                        instRow.Args = relayFunction.Function.ArgList;
                        InstanceRow existItem = commonSheet.Rows.Find(x => x.TestName.Equals(instRow.TestName, StringComparison.OrdinalIgnoreCase));
                        if (existItem != null)
                        {
                            commonSheet.Rows.Remove(existItem);
                        }

                        commonSheet.AddRow(instRow);
                    }
                }
            }

            Response.Report("Parsing Relay Table Done!", EnumMessageLevel.General, 50);
            return sheet.RelayItems;
        }
    }
}
