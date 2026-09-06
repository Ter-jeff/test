using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using OfficeOpenXml;

namespace Automation.InputManager
{
    public abstract class InputManagerBase<T> where T : InputDataBase
    {
        protected readonly ExcelWorkbook ExcelWorkbook;
        protected ParaData ParaData;

        protected InputManagerBase(ExcelWorkbook excelWorkbook)
        {
            ExcelWorkbook = excelWorkbook;
        }

        public abstract T Read();
    }
}
