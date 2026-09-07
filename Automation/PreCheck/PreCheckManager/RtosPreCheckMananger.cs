using Automation.PreCheck.AllParaData;

using OfficeOpenXml;

namespace Automation.PreCheck.PreCheckManager
{
    public class RtosPreCheckManager : PreCheckManager<ParaData>
    {
        public RtosPreCheckManager(ExcelWorkbook excelWorkbook, ParaData paraData) : base(excelWorkbook, paraData)
        {
        }

        public override bool PreCheckAll()
        {
            return false;
        }
    }
}
