using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;

using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Rtos
{
    [TestClass]
    public class RtosBinTableRowGeneratorTests
    {
        [TestMethod]
        public void GenBinTableRow_IdsCurrent_With_IdsNoFuse_Should_Return_FailAction()
        {
            var generator = new TestableRtosBinTableRowGeneratorTests(
                sheetName: "RTOS",
                errorBinNums: []
            );

            var pattern = new HardIpPattern
            {
                FunctionName = FuncNameConst.CSharpFuncNameIdsCurrent,
                Failflag = "",
                MiscInfo = "IDS_NO_FUSE"
            };

            generator.SetPattern_Public(pattern);

            BinTableRow row = generator.GenBinTableRow_Public();

            Assert.AreNotEqual(null, row);
            Assert.IsTrue(row.ItemList.Contains(HardIpConstData.PrefixHardIpFailAction));
        }

        [TestMethod]
        public void GenBinTableRow_When_FunctionName_Is_Empty_Should_Use_Default_BinParameter()
        {
            var generator = new TestableRtosBinTableRowGeneratorTests(
                sheetName: "RTOS",
                errorBinNums: []
            );

            var pattern = new HardIpPattern
            {
                FunctionName = "",
                Failflag = "",
                MiscInfo = ""
            };

            generator.SetPattern_Public(pattern);

            BinTableRow row = generator.GenBinTableRow_Public();

            Assert.AreNotEqual(null, row);
            Assert.IsFalse(string.IsNullOrEmpty(row.ItemList));
        }
    }
}
