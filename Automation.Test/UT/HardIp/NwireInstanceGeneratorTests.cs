using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class NwireInstanceGeneratorTests : FunctionTestBase
    {
        private static NwireInstanceGenerator _generator = null!;
        private static HardIpSheet _hardIpSheet = null!;
        private static InstanceSheet _instanceSheet = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            LocalSpecs.TarFolder = OutputPath;

            _generator = new NwireInstanceGenerator();

            TestProgram.Clear();
            var commonSheet = new InstanceSheet(SheetConst.TestInstCommon);
            TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirCommonSheets, commonSheet);

            _hardIpSheet = new HardIpSheet();
            _hardIpSheet.Rows.Add(new HardIpPattern
            {
                Pattern = new PatternClass("PAT_A"),
                SheetName = "HARDIP_BLOCK1",
                ForceCondition = new ForceClass { ForceCondition = "AC:RT_CLK32768:10Mhz" }
            });

            _instanceSheet = new InstanceSheet("InstSheet1");
            var instRow = new InstanceRow
            {
                Args = ["PAT_A"]
            };
            _instanceSheet.AddRow(instRow);
        }

        [TestMethod]
        public void GenNwire_ShouldAddNwireInstanceRows()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet>
            {
                { "HardIP_Block1", _hardIpSheet }
            };
            var instSheets = new List<InstanceSheet> { _instanceSheet };

            // Act
            _generator.GenNwireInstance(planDic, instSheets);

            // Assert
            InstanceSheet? commonSheet = TestProgram.IgxlWorkBk.InsSheets.Values.FirstOrDefault(s => s.Name == SheetConst.TestInstCommon);
            Assert.AreNotEqual(null, commonSheet);
            Assert.IsTrue(commonSheet!.Rows.Count == 2);
        }

        [TestMethod]
        public void GetAcCategoryByAc_ShouldReturnExpectedCategory()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                SheetName = "My_REG_DC_TEST_IDS_Sheet",
                SubBlock = "CP",
                ForceCondition = new ForceClass { ForceCondition = "AC:RT_CLK32768:10Mhz" }
            };
            string result = _generator.GetAcCategoryByAc(pattern, "BLOCKA", "TimeSet1");

            // Assert
            Assert.AreEqual("BLOCKA_TBDRT_CLK32768_10Mhz", result);
        }
    }
}
