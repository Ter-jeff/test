using Automation.PreCheck.PreChecks.Basic;

using OfficeOpenXml;

namespace Automation.PreCheck.PreChecks.PreAction
{
    public abstract class PreActionCheckBase : PreCheckBase
    {
        protected PreActionCheckBase(ExcelWorkbook excelWorkbook, string sheetName) : base(excelWorkbook, sheetName)
        {
        }
    }
}
