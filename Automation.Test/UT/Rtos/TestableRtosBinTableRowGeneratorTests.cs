using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow;
using Automation.GenerateIgxl.HardIp.InputObject;

using IgxlLib.IgxlBase;

namespace Automation.Test.UT.Rtos
{
    internal class TestableRtosBinTableRowGeneratorTests(string sheetName, List<string> errorBinNums) : RtosBinTableRowGenerator(sheetName, errorBinNums)
    {
        public void SetPattern_Public(HardIpPattern hardIpPattern)
        {
            base.SetPattern(hardIpPattern);
        }

        public void OverrideFunctionName(string functionName)
        {
            // ⭐ 關鍵：直接改 generator 內部 Pattern
            Pattern.FunctionName = functionName;
        }

        public void OverrideFailFlag(string failFlag)
        {
            Pattern.Failflag = failFlag;
        }

        public BinTableRow GenBinTableRow_Public()
        {
            return base.GenBinTableRowForPattern();
        }
    }
}
