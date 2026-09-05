using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow
{
    public class FreqFlowSheetGenerator : HardIpFlowSheetGenerator
    {
        public FreqFlowSheetGenerator(HardIpInputData hardIpInputData, string sheetName, List<HardIpPattern> patternList = null) : base(hardIpInputData, sheetName, patternList)
        {
            FlowRowGenerator = new FreqFlowRowGenerator(hardIpInputData, sheetName);
        }
    }
}
