using System.Collections.Generic;

using Automation.GenerateIgxl.BinCut.Business;
using Automation.InputManager.Data;
using Automation.Static;
using Automation.Test.Static;

using CommonLib.Extension;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using TestPlanLib.BinCut.Binning;
using TestPlanLib.VbtLib;

namespace Automation.Test.UT.BinCut
{
    [TestClass]
    public class BinCutFlowInstanceWriterCsTests
    {
        private Mock<VbtFunctionLib> _mockLib = null!;
        private BinCutInputData _inputData = null!;

        [TestInitialize]
        public void Setup()
        {
            _mockLib = new Mock<VbtFunctionLib>();
            TestProgram.VbtFunctionLib = _mockLib.Object;
            _inputData = new BinCutInputData();
            _inputData.BinningTables.Add(new BinningTable());
        }

        [TestCleanup]
        public void ClassCleanup()
        {
            var vbtFunctionLib = new VbtFunctionLib();
            TestProgram.VbtFunctionLib = vbtFunctionLib;
            TestProgram.VbtFunctionLib.AddVbtFunctionRange(TestSuiteInitialize.Functions);
        }

        [TestMethod]
        public void GenPrintOutVddBinning_ReturnsDotNetInstance_WhenFunctionFound()
        {
            // Arrange
            var mockFunction = new Function
            {
                IsFound = false,
                FunctionName = "Judge_Stored_IDS_Func",
                Parameters = "Param1,Param2"
            };
            _mockLib.Setup(l => l.GetFunctionByName("PrintVddBinning", "bincut", true)).Returns(mockFunction);

            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);

            // Act
            IgxlLib.IgxlBase.InstanceRow result = writer.GenPrintOutVddBinning();

            // Assert
            Assert.AreEqual("PrintOut_VDD_Bin", result.VbtName);
            Assert.AreEqual("PrintOutVddBinning", result.TestName);
        }

        [TestMethod]
        public void GenAdjustVddBinningInstance()
        {
            // Arrange
            var mockFunction = new Function
            {
                IsFound = false,
                FunctionName = "AdjustVddBinning",
                Parameters = "Param1,Param2"
            };
            _mockLib.Setup(l => l.GetFunctionByName(It.Is<string>(s => s.EqualsIgnoreCase("AdjustVddBinning")), It.Is<string>(s => s.EqualsIgnoreCase("bincut")), It.IsAny<bool>())).Returns(mockFunction);

            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);

            // Act
            string testArgument = "AdjustVddBinning";
            IgxlLib.IgxlBase.InstanceRow result = writer.GenAdjustVddBinningInstance(testArgument);

            // Assert
            Assert.AreEqual("AdjustVddBinning", result.VbtName);
            Assert.AreEqual("AdjustVddBinning", result.TestName);
        }

        [TestMethod]
        [DataRow("Max PV (VDD_SOC)", "AdjustVddBinning")]
        [DataRow("Min PV (VDD_CORE)", "AdjustVddBinning")]
        [DataRow("", "AdjustVddBinning")]
        public void GenAdjustVddBinningInstance_DataDriven(string text, string expected)
        {
            // 1. Arrange - Setup Function
            var mockFunction = new Function
            {
                IsFound = true,
                Type = ".NET",
                FunctionName = "Bincut.AdjustVddBinning",
                Parameters = "Param1,Param2"
            };

            _mockLib.Setup(l => l.GetFunctionByName(It.IsAny<string>(), It.Is<string>(s => s.EqualsIgnoreCase("bincut")), It.IsAny<bool>()))
                .Returns(mockFunction);

            // 2. Arrange - Setup BinningTable with RowData
            var table = new BinningTable
            {
                CommentIdx = 0,

                // Create the row data using the 'text' parameter from DataRow
                Rows =
                [
                    new() { RowData = [text] }
                ]
            };

            // Initialize the list to avoid First() exceptions
            _inputData.BinningTables = [table];

            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);

            // 3. Act
            string testArgument = "AdjustVddBinning";
            IgxlLib.IgxlBase.InstanceRow result = writer.GenAdjustVddBinningInstance(testArgument);

            // 4. Assert
            Assert.AreNotEqual(null, result, "Resulting InstanceRow should not be null");
            Assert.AreEqual(expected, result.TestName, $"Failed for input: {text}");
            Assert.AreEqual(".Bincut.AdjustVddBinning", result.VbtName);
        }

        [TestMethod]
        public void GenAdjustVddBinningInstance_1()
        {
            // Arrange
            var mockFunction = new Function
            {
                IsFound = true,
                Type = ".NET",
                FunctionName = "Bincut.AdjustVddBinning",
                Parameters = "Param1,Param2"
            };

            _mockLib.Setup(l => l.GetFunctionByName(It.IsAny<string>(), It.Is<string>(s => s.EqualsIgnoreCase("bincut")), It.IsAny<bool>())).Returns(mockFunction);

            _inputData.BinningTables.Clear();

            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);

            // 3. Act
            string testArgument = "AdjustVddBinning";
            IgxlLib.IgxlBase.InstanceRow result = writer.GenAdjustVddBinningInstance(testArgument);

            // 4. Assert
            Assert.AreNotEqual(null, result, "Resulting InstanceRow should not be null");
            Assert.AreEqual(".Bincut.AdjustVddBinning", result.VbtName);
        }

        [TestMethod]
        public void GenCheck_IDS()
        {
            // Arrange
            var mockFunction = new Function
            {
                IsFound = false,
                FunctionName = "Judge_Stored_IDS_Func",
                Parameters = "Param1,Param2"
            };
            _mockLib.Setup(l => l.GetFunctionByName("Judge_Stored_IDS", "bincut", true)).Returns(mockFunction);

            // Act
            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);
            IgxlLib.IgxlBase.InstanceRow result = writer.GenCheck_IDS();

            // Assert
            Assert.AreEqual("check_IDS", result.VbtName);
        }

        [TestMethod]
        public void GenPower_Binning_Calculation()
        {
            // Arrange
            var mockFunction = new Function
            {
                IsFound = false,
                FunctionName = "PowerBinning",
                Parameters = "Param1,Param2"
            };
            _mockLib.Setup(l => l.GetFunctionByName("PowerBinning", "bincut", true)).Returns(mockFunction);

            // Act
            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);
            IgxlLib.IgxlBase.InstanceRow result = writer.GenPower_Binning_Calculation();

            // Assert
            Assert.AreEqual("Power_Binning_Calculation", result.VbtName);
        }

        [TestMethod]
        public void GenPrintConfigInstanceRow()
        {
            // Arrange
            var mockFunction = new Function
            {
                IsFound = false,
                FunctionName = "PrintOutBinCutConfig",
                Parameters = "Param1,Param2"
            };
            _mockLib.Setup(l => l.GetFunctionByName("PrintOutBinCutConfig", "bincut", true)).Returns(mockFunction);
            _mockLib.Setup(l => l.GetFunctionByName("Print_BinCut_config", "bincut", It.IsAny<bool>())).Returns(mockFunction);

            // Act
            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);
            var testArgument = new List<string>();
            IgxlLib.IgxlBase.InstanceRow result = writer.GenPrintConfigInstanceRow(testArgument);

            // Assert
            Assert.AreEqual("PrintOutBinCutConfig", result.VbtName);
        }

        [TestMethod]
        public void GenSet_VBinResult_without_Test()
        {
            // Arrange
            var mockFunction = new Function
            {
                IsFound = false,
                FunctionName = "SetVoltageWithoutTest",
                Parameters = "Param1,Param2"
            };
            _mockLib.Setup(l => l.GetFunctionByName("SetVoltageWithoutTest", "bincut", true)).Returns(mockFunction);

            // Act
            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);
            var testArgument = new Dictionary<string, bool>();
            IgxlLib.IgxlBase.InstanceRow result = writer.GenSet_VBinResult_without_Test(testArgument);

            // Assert
            Assert.AreEqual("Set_VBinResult_without_Test", result.VbtName);
        }

        [TestMethod]
        public void GenFuseBinnedProductVoltagesInstanceRow()
        {
            // Arrange
            var mockFunction = new Function
            {
                IsFound = false,
                FunctionName = "FuseBinCutResults",
                Parameters = "Param1,Param2"
            };
            _mockLib.Setup(l => l.GetFunctionByName("FuseBinCutResults", "bincut", true)).Returns(mockFunction);

            // Act
            var writer = new BinCutFlowInstanceWriterCs(_inputData, []);
            IgxlLib.IgxlBase.InstanceRow? result = writer.GenFuseBinnedProductVoltagesInstanceRow();

            // Assert
            Assert.AreEqual(null, result);
        }
    }
}
