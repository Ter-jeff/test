using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTable
{
    public class HardIpBinTableGenerator : BlockBinTableGeneratorBase
    {
        public HardIpBinTableGenerator(HardIpInputData hardIpInputData, BinTableSheet hardIpBinTableSheet, string sheetName, List<HardIpPattern> patternList, List<string> duplicateParameter, List<string> errorBinNums)
            : base(hardIpInputData, hardIpBinTableSheet, patternList, duplicateParameter)
        {
            BinTableRowGenerator = new HardIpBinTableRowGenerator(sheetName, errorBinNums);
        }
    }
}
