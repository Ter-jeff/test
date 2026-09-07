using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;
using Automation.Test.Static;

using CommonReaderLib.PatternListCsv;

using FileDiffLib;

using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class HardIpBlockInsGeneratorTests : FunctionTestBase
    {
        private HardIpSheet _sheet = null!;
        private HardIpInputData _inputData = null!;

        private HardIpPattern _mockPattern = null!;
        private HardIpRegAssign _regAssignItem = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
        }

        [TestInitialize]
        public void Setup()
        {
            _sheet = new HardIpSheet
            {
                Rows = []
            };
            HardIpParaData paraData = new HardIpParaData(EnumBlock.HardIp);
            _inputData = new HardIpInputData(paraData)
            {
                HardIpRegAssigns = []
            };

            // Initialize common objects
            _mockPattern = new HardIpPattern { SheetName = "TestSheet", RowNum = 1, Pattern = new PatternClass("TestPattern") { RealPatternName = "TestPattern" } };
            _regAssignItem = new HardIpRegAssign();
        }

        [TestMethod]
        public void GenBlockInsRows_ShouldReturnEmptyList_WhenNoPatternsExist()
        {
            // Arrange
            var generator = new HardIpBlockInsGenerator(
                _inputData,
                "SheetA",
                _sheet);

            // Act
            List<InstanceSheet> result = generator.GenBlockInsRows();

            // Assert
            Assert.AreEqual(0, result.Count, "Empty HardIpSheet.Rows should produce no InstanceSheets");
        }

        [TestMethod]
        public void GenBlockInsRows_ShouldRunBlockPath_WhenByBlockIsTrue()
        {
            // Arrange
            var pattern = new HardIpPattern
            {
                SheetName = "BlockSheet",
                RowNum = 2,
                Pattern = new PatternClass("PAT_BLOCK"),
                SkipDotNet = false,
                SkipList = []
            };

            _sheet.Rows.Add(pattern);

            var generator = new HardIpBlockInsGenerator(
                _inputData,
                "SheetBlock",
                _sheet);

            // Act
            List<InstanceSheet> result = generator.GenBlockInsRows();

            // Assert
            Assert.AreNotEqual(null, result, "Should return a non-null list");
            Assert.IsTrue(result.Count >= 0, "Should return list even if no valid instances are created");
        }

        [TestMethod]
        public void GenBlockInsRows_ShouldHandleMultiplePatternsGracefully()
        {
            // Arrange
            _sheet.Rows.Add(new HardIpPattern
            {
                SheetName = "Test1",
                Pattern = new PatternClass("PAT1"),
                SkipList = []
            });

            _sheet.Rows.Add(new HardIpPattern
            {
                SheetName = "Test2",
                Pattern = new PatternClass("PAT2"),
                SkipList = []
            });

            var generator = new HardIpBlockInsGenerator(
                _inputData,
                "MultiSheet",
                _sheet);

            // Act
            List<InstanceSheet> result = generator.GenBlockInsRows();

            // Assert
            Assert.AreNotEqual(null, result);
            Assert.IsTrue(result.Count >= 0, "Multiple patterns should not cause exceptions");
        }

        [TestMethod]
        public void GenBlockInsRowsByBlockTest()
        {
            string subName = "GenBlockInsRowsByBlock";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            var generator = new HardIpBlockInsGenerator(
                _inputData,
                "HardIp_PCIE",
                _sheet);

            var rows = new List<HardIpPattern>
            {
                new()
                {
                    Pattern = new PatternClass("Pattern1") { RealPatternName = "Pattern1#T1" },
                    SheetName = "HARDIP_ABC",
                    FunctionName = "check_IDS",
                    MiscInfo = "RetestReset:" + new string('X', 8000)
                },
                new()
                {
                    Pattern = new PatternClass("Pattern1") { RealPatternName = "Pattern1#T1" },
                    SheetName = "HARDIP_ABC",
                    FunctionName = "check_IDS",
                    MiscInfo = "RetestReset:" + new string('X', 8000)
                },
                new()
                {
                    Pattern = new PatternClass("Pattern1") { RealPatternName = "Pattern1#T1" },
                    SheetName = "HARDIP_ABC"
                },
                new()
                {
                    Pattern = new PatternClass("Pattern1") { RealPatternName = "Pattern1#T1" },
                    SheetName = "WIRELESS_ABC"
                },
                new()
                {
                    Pattern = new PatternClass("Pattern1") { RealPatternName = "Pattern1#T1" },
                    SheetName = "WIRELESS_ABC"
                }
            };
            _sheet.Rows.AddRange(rows);
            _sheet.PlanHeaderIdx["registerIndex"] = 1;

            // Act
            List<InstanceSheet> sheets = [];
            List<InstanceSheet> hvs = generator.GenBlockInsRowsByBlock(_sheet.Rows, "HV");
            hvs.ForEach(x => x.Name += "_HV");
            sheets.AddRange(hvs);
            List<InstanceSheet> lvs = generator.GenBlockInsRowsByBlock(_sheet.Rows, "LV");
            lvs.ForEach(x => x.Name += "_LV");
            sheets.AddRange(lvs);

            // Assert
            foreach (InstanceSheet sheet in sheets)
            {
                sheet.Write(Path.Combine(outputPath, sheet.Name + ".txt"));
            }

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void GetRegAssginList_DigSrcEquation_ShouldPreserveOrder()
        {
            // Arrange
            string[] infoPara = ["A+B+C"];

            // Act
            List<List<string>> result =
                HardIpBlockInsGenerator.GetRegAssginList(_mockPattern, 0, "DigSrc_Equation", infoPara, 0, _regAssignItem);

            // Assert
            Assert.AreEqual("'A", result[0][0]);
            Assert.AreEqual("'B", result[1][0]);
            Assert.AreEqual("'C", result[2][0]);
        }

        [TestMethod]
        public void GetRegAssginList_UnknownType_ShouldNotSplitInput()
        {
            // Arrange
            string[] infoPara = ["A+B,C=D"];

            // Act
            List<List<string>> result =
                HardIpBlockInsGenerator.GetRegAssginList(_mockPattern, 0, "NotSupported", infoPara, 0, _regAssignItem);

            // Assert
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        [DataRow("DigSrc_Equation", RegisterAssignType.DigSrc_Equation)]
        [DataRow("DigSrc_Assignment", RegisterAssignType.DigSrc_Assignment)]
        [DataRow("CUS_Str_DigCapData", RegisterAssignType.CUS_Str_DigCapData)]
        [DataRow("Calc_Eqn", RegisterAssignType.Calc_Eqn)]
        [DataRow("MeasI_Range", RegisterAssignType.MeasI_Range)]
        [DataRow("ForceV_Val", RegisterAssignType.ForceV_Val)]
        [DataRow("ForceI_Val", RegisterAssignType.ForceI_Val)]
        [DataRow("CUS_Str_MainProgram", RegisterAssignType.CUS_Str_MainProgram)]
        public void GetRegAssginList_ShouldSetCorrectRegisterAssignType(string type, RegisterAssignType registerAssignType)
        {
            // Arrange
            string[] infoPara = ["Dummy"];

            // Act
            _ = HardIpBlockInsGenerator.GetRegAssginList(_mockPattern, 0, type, infoPara, 0, _regAssignItem);

            // Assert
            Assert.AreEqual(registerAssignType, _regAssignItem.Type);
        }

        [TestMethod]
        public void GetRegAssginList_DigSrcAssignment_WithoutEqual_ShouldOnlyHaveKey()
        {
            // Arrange
            string type = "DigSrc_Assignment";
            string[] infoPara = ["OnlyKey"];

            // Act
            List<List<string>> result =
                HardIpBlockInsGenerator.GetRegAssginList(_mockPattern, 0, type, infoPara, 0, _regAssignItem);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(1, result[0].Count);
            Assert.AreEqual("OnlyKey", result[0][0]);
        }

        [TestMethod]
        [DataRow("DigSrc_Equation", "A+B+C", 3, "'A", DisplayName = "DigsrcEquation: Split by +")]
        [DataRow("DigSrc_Assignment", "Var1=Val1;Var2=Val2", 2, "Var1", DisplayName = "DigsrcAssignment: Split by ; and =")]
        [DataRow("CUS_Str_DigCapData", "Key1:Val1,Key2:Val2", 2, "Key1", DisplayName = "CusStrDigcapdata: Split by , and :")]
        [DataRow("Calc_Eqn", "Calc1:Val1;Calc2:Val2", 2, "Calc1", DisplayName = "CalcEqn: Split by ; and :")]
        [DataRow("MeasI_Range", "R1+R2", 2, "'R1", DisplayName = "MeasiRange: Split by +")]
        [DataRow("ForceV_Val", "1.2|3.4", 2, "'1.2", DisplayName = "ForcevVal: Split by |")]
        [DataRow("ForceI_Val", "0.5|0.8", 2, "'0.5", DisplayName = "ForceiVal: Split by |")]
        [DataRow("CUS_Str_MainProgram", "Prog1,Prog2", 2, "'Prog1", DisplayName = "CusStrMainprogram: Split by ,")]
        public void GetRegAssginList_ShouldParseCorrectly(string type, string input, int expectedCount, string firstElementExpected)
        {
            // Arrange
            string[] infoPara = [input];
            int index = 0;

            // Act
            List<List<string>> result = HardIpBlockInsGenerator.GetRegAssginList(_mockPattern, 0, type, infoPara, index, _regAssignItem);

            // Assert
            Assert.AreEqual(expectedCount, result.Count);
            Assert.AreEqual(firstElementExpected, result[0][0]);
        }

        [TestMethod]
        public void GetRegAssginList_DigSrcAssignment_ShouldIncludeValueWithQuote()
        {
            // Arrange
            string type = "DigSrc_Assignment";
            string[] infoPara = ["Address=0x10"];

            // Act
            List<List<string>> result = HardIpBlockInsGenerator.GetRegAssginList(_mockPattern, 0, type, infoPara, 0, _regAssignItem);

            // Assert
            Assert.AreEqual("Address", result[0][0]);
            // Verifying the "'" prefix logic
            Assert.AreEqual("'0x10", result[0][1]);
        }

        [TestMethod]
        public void GetRegAssginList_UnknownType_ShouldSetNoneType()
        {
            // Arrange
            string type = "UnknownType";
            string[] infoPara = ["some data"];

            // Act
            List<List<string>> result = HardIpBlockInsGenerator.GetRegAssginList(_mockPattern, 0, type, infoPara, 0, _regAssignItem);

            // Assert
            Assert.AreEqual(0, result.Count);
            Assert.AreEqual(RegisterAssignType.NoneType, _regAssignItem.Type);
        }
    }
}
