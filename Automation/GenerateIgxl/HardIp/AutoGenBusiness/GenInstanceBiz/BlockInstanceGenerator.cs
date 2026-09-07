using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public abstract class BlockInstanceGenerator
    {
        protected HardIpInputData HardIpInputData { get; }
        protected string SheetName = string.Empty;
        protected HardIpSheet HardIpSheet;
        protected InsRowGenerator InstanceRowGenerator = null;

        protected BlockInstanceGenerator(HardIpInputData hardIpInputData, string sheetName, HardIpSheet hardIpSheet)
        {
            HardIpInputData = hardIpInputData;
            SheetName = sheetName;
            HardIpSheet = hardIpSheet;
        }

        public abstract List<InstanceSheet> GenBlockInsRows();
    }
}
