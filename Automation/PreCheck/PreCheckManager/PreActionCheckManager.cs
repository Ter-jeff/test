using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using OfficeOpenXml;

namespace Automation.PreCheck.PreCheckManager
{
    public class PreActionCheckManager : PreCheckManager<ParaData>
    {
        private readonly PreActionInputData _preActionInputData;

        public PreActionCheckManager(ExcelWorkbook excelWorkbook, ParaData paraData, PreActionInputData preActionInputData) : base(excelWorkbook, paraData)
        {
            _preActionInputData = preActionInputData;
        }

        public override bool PreCheckAll()
        {
            if (_preActionInputData.IoContinuity != null && _preActionInputData.PinMapSheet != null)
            {
                _preActionInputData.IoContinuity.CheckIoPins(_preActionInputData.PinMapSheet);
            }

            return false;
        }
    }
}
