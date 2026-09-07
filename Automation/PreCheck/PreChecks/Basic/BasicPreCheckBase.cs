using OfficeOpenXml;

namespace Automation.PreCheck.PreChecks.Basic
{
    public abstract class BasicPreCheckBase : PreCheckBase
    {
        protected BasicPreCheckBase(ExcelWorkbook excelWorkbook, string sheetName) : base(excelWorkbook, sheetName)
        {
        }
    }
}
