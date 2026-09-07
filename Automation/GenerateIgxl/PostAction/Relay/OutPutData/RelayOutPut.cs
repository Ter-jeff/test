using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.Relay.IGXLData;
using Automation.GenerateIgxl.PostAction.Relay.InPutDataStruct;
using Automation.GenerateIgxl.PostAction.Relay.RelayConst;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.Relay.OutPutData
{
    public class RelayOutPut
    {
        public Dictionary<string, RelayFlowRow> RelayPadFlowRowList = new Dictionary<string, RelayFlowRow>();
        public RelayInstanceSheet RelayInstance = new RelayInstanceSheet(OutputConst.RelayInstName);
        public List<RelayInputData> InputData = new List<RelayInputData>();

        public void PadDataToFlowSheet()
        {
            foreach (KeyValuePair<string, RelayFlowRow> relayPadFlowRow in RelayPadFlowRowList)
            {
                foreach (KeyValuePair<string, SubFlowSheet> flowSheetPair in TestProgram.IgxlWorkBk.SubFlowSheets)
                {
                    if (flowSheetPair.Value.Name == relayPadFlowRow.Key)
                    {
                        flowSheetPair.Value.InsertRow(relayPadFlowRow.Value.InsertIndex, relayPadFlowRow.Value);
                    }
                }
            }
        }

        public void AddInstanceSheetToLocalSpace()
        {
            bool isExist = false;
            foreach (KeyValuePair<string, InstanceSheet> insSheetPair in TestProgram.IgxlWorkBk.InsSheets)
            {
                if (insSheetPair.Value.Name == RelayInstance.Name)
                {
                    isExist = true;
                    foreach (InstanceRow insRow in RelayInstance.Rows)
                    {
                        insSheetPair.Value.AddRow(insRow);
                    }
                }
            }
            if (!isExist)
            {
                string commonSheet = TestProgram.IgxlWorkBk.InsSheets.Keys.ToList().FirstOrDefault(x => x.Contains(Path.DirectorySeparatorChar + "TestInst_Common"));
                if (commonSheet != null && TestProgram.IgxlWorkBk.InsSheets.ContainsKey(commonSheet))
                {
                    TestProgram.IgxlWorkBk.InsSheets[commonSheet].AddRows(RelayInstance.Rows);
                }
                else
                {
                    TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, RelayInstance);
                }
            }
        }
    }
}
