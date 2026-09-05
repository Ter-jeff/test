using Automation.Static;

using IgxlLib;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.EndFlow
{
    public class MainEndFlow
    {
        private SubFlowSheet _mainEndFlowSheet;

        public void WorkFlow()
        {
            _mainEndFlowSheet = TestProgram.IgxlWorkBk.GetFlowSheet(IgxlWorkBook.FlowTableMainEndFlow, FolderStructure.DirMain, "END");

            _mainEndFlowSheet.AddReturnRow();

            TestProgram.IgxlWorkBk.SetFlowSheet(IgxlWorkBook.FlowTableMainEndFlow, _mainEndFlowSheet);
        }
    }
}
