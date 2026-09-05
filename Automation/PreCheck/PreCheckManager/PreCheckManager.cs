using Automation.PreCheck.AllParaData;

using OfficeOpenXml;

namespace Automation.PreCheck.PreCheckManager
{
    public abstract class PreCheckManager<T1> where T1 : ParaData
    {
        protected ExcelWorkbook ExcelWorkbook;
        protected readonly T1 ParaData;

        protected PreCheckManager(ExcelWorkbook excelWorkbook, T1 paraData)
        {
            ExcelWorkbook = excelWorkbook;
            ParaData = paraData;
        }

        public abstract bool PreCheckAll();
    }
}
