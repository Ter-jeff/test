using System.Collections.Generic;

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
    public class AcCategoryGeneratorTests : FunctionTestBase
    {
        private AcCategoryGenerator _generator = null!;
        private AcSpecSheet _acSpecSheet = null!;

        [TestInitialize]
        public void Setup()
        {
            _generator = new AcCategoryGenerator();

            _acSpecSheet = new AcSpecSheet("TestSheet");
            _acSpecSheet.Rows.Add(new AcSpec("PIN1_Freq_VAR", [new("Typ", "Typ"), new("Min", "Min"), new("Max", "Max")]));
            _acSpecSheet.AddCategoryList("DEFAULT_CAT");

            TestProgram.Clear();
            TestProgram.IgxlWorkBk.AcSpecSheets["TestSheet"] = _acSpecSheet;
        }

        [TestMethod]
        public void GenAcCategory_ShouldAddCategories()
        {
            // Arrange
            var planDic = new Dictionary<string, HardIpSheet>();
            var hardIpSheet = new HardIpSheet();
            hardIpSheet.Rows.Add(new HardIpPattern
            {
                Pattern = new PatternClass("PAT1")
                {
                    PatternSetList = [["P1"]]
                },
                SheetName = "HardIP_Test",
                ForceCondition = new ForceClass { ForceCondition = "AC:TCK:12Mhz" }
            });
            planDic.Add("Sheet1", hardIpSheet);

            var instSheet = new InstanceSheet("InstSheet");
            var row = new InstanceRow
            {
                Args = ["P1"]
            };
            instSheet.Rows.Add(row);
            var instSheets = new List<InstanceSheet> { instSheet };

            // Act
            _generator.GenAcCategory(planDic, instSheets);

            // Assert
            Assert.IsTrue(_acSpecSheet.CategoryList.Count == 2);
        }
    }
}
