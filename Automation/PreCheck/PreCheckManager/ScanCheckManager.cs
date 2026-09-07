using Automation.PreCheck.AllParaData;

using OfficeOpenXml;

namespace Automation.PreCheck.PreCheckManager
{
    public class ScanCheckManager : PreCheckManager<ParaData>
    {
        public ScanCheckManager(ExcelWorkbook excelWorkbook, ParaData paraData) : base(excelWorkbook, paraData)
        {
        }

        public override bool PreCheckAll()
        {
            new HarvestGroupPinsChecker().Check();

            return false;
        }
    }
}
