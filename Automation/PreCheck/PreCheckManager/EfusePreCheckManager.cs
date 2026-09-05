using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.InputChecker;
using Automation.InputManager.Data;

namespace Automation.PreCheck.PreCheckManager
{
    public class EfusePreCheckManager
    {
        public void Check(EFuseInputData eFuseInputData)
        {
            if (eFuseInputData.EfuseConfigMainSheets.Any() && eFuseInputData.EfuseBitDefTables.Any() && eFuseInputData.EfuseBitDefTables.Exists(x => EFuseConst.GetBankName(x.BlockName).Equals(BankType.Cfg)))
            {
                new EfuseConfigureAllChecker().WorkFlow(eFuseInputData.EfuseConfigMainSheets, eFuseInputData.EfuseBitDefTables);
                foreach (TestPlanLib.Efuse.Input.EfuseConfigMainSheet efuseConfigMainSheet in eFuseInputData.EfuseConfigMainSheets)
                {
                    efuseConfigMainSheet.AddToErrorReport();
                }
                new EfuseConfigureBdfChecker().WorkFlow(eFuseInputData.EfuseConfigMainSheets.First(), eFuseInputData.EfuseBitDefTables, eFuseInputData.EfusePatternRows);
            }
        }
    }
}
