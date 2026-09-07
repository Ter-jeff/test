using Automation.Static;

using IgxlLib;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PreAction.GenEmptyBinTbl
{
    public class EmptyBinTblMain
    {
        public void WorkFlow()
        {
            BinTableSheet emptyBinTblSheet = new BinTableSheet(IgxlWorkBook.MainBinTblSheetName);
            TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirBinTable, emptyBinTblSheet);
        }
    }
}
