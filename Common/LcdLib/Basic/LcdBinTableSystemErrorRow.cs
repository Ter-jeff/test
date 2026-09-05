using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinNumberLegacy;
using TestPlanLib.Singleton;

namespace LcdLib.Basic
{
    internal static class LcdBinTableSystemErrorRow
    {
        internal static void Add()
        {
            BinTableSheet binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);

            var para = new BinNumDefPara(EnumBinNumDefBlock.SystemError, "system error");
            BinNumberSingletonLegacy.Instance().GetBinNumDefRow(para, out BinNumDefRow bin);
            var binRow = new BinTableRow
            {
                Name = "Bin_System_Error",
                Op = "AND",
                Sort = bin.CurrentSoftBin.ToString(),
                Bin = bin.HardBin,
                Result = "Fail"
            };
            binRow.Items.Add("T");
            binTable.AddRow(binRow);
        }
    }
}
