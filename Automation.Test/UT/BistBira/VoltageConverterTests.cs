using Automation.GenerateIgxl.BistBira.BistInputLib;
using Automation.Singleton;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using ScghLib.Reader;

using TestPlanLib.DataStruct;

namespace Automation.Test.UT.BistBira
{
    [TestClass]
    public class VoltageConverterTests
    {
        private Mock<MultiTestSettingSheetsSingleton> _mockSetting = null!;
        private VoltageConverter _converter = null!;

        [TestInitialize]
        public void Setup()
        {
            // Note: If MultiTestSettingSheetsSingleton does not have a parameterless constructor 
            // or virtual methods, you may need to wrap it in an Interface.
            _mockSetting = new Mock<MultiTestSettingSheetsSingleton>();
            _converter = new VoltageConverter(_mockSetting.Object);
        }

        [TestMethod]
        public void WorkFlow_SkipsRow_WhenRetentionTypeIsEmpty()
        {
            // Arrange
            BistProdFlowSheet prodFlow = CreateMockFlow("P1_ERT_2", "NV");

            // Act
            _converter.WorkFlow(ref prodFlow);

            // Assert
            Assert.AreEqual("NV", prodFlow.Rows[0].Voltage);
        }

        [TestMethod]
        public void WorkFlow_UpdatesVoltage_WhenCategoryExistsInMultiTest()
        {
            // Arrange
            var row = new BistProdFlowRow
            {
                Pattern = "P1_ERT_2", // Triggers CheckRetention
                Voltage = "ExistingCat, NV",
                VoltageMode = "High"
            };
            var prodFlow = new BistProdFlowSheet { Rows = [row] };
            prodFlow.MbistSheet.SheetName = "Module_A0";

            _mockSetting.Object.DcCategoryInfos = [new DcCategoryInfo("ExistingCat")];
            _mockSetting.Setup(m => m.GetDcCategoryVoltages("ExistingCat")).Returns(["HV", "LV"]);

            // Act
            _converter.WorkFlow(ref prodFlow);

            // Assert
            Assert.AreEqual("ExistingCat,HV", row.Voltage);
        }

        [TestMethod]
        public void WorkFlow_SetsDefaultVoltage_WhenCategoryDoesNotExist()
        {
            // Arrange
            var row = new BistProdFlowRow
            {
                Pattern = "P1_ERT_2",
                Voltage = "NV",
                VoltageMode = "Nominal"
            };
            var prodFlow = new BistProdFlowSheet { Rows = [row] };
            prodFlow.MbistSheet.SheetName = "Module_A0";
            _mockSetting.Object.DcCategoryInfos = [];
            _mockSetting.Setup(m => m.FindMbistCatgeoryName(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), null, out It.Ref<EnumMessageLevel>.IsAny, out It.Ref<string>.IsAny, "A0", It.IsAny<string>())).Returns("FoundCategory");
            _mockSetting.Setup(m => m.GetDcCategoryVoltages("FoundCategory")).Returns(["NV"]);

            // Act
            _converter.WorkFlow(ref prodFlow);

            // Assert
            Assert.AreEqual("NV", row.Voltage);
        }

        [TestMethod]
        public void WorkFlow_SetsDefaultVoltage_WhenCategoryDoesNotExist_1()
        {
            // Arrange
            var row = new BistProdFlowRow
            {
                Pattern = "P1_ERT_2",
                Voltage = "NV",
                VoltageMode = "Nominal"
            };
            var prodFlow = new BistProdFlowSheet { Rows = [row] };
            prodFlow.MbistSheet.SheetName = "Module_A0";
            _mockSetting.Object.DcCategoryInfos = [];
            _mockSetting.Setup(m => m.FindMbistCatgeoryName(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), null, out It.Ref<EnumMessageLevel>.IsAny, out It.Ref<string>.IsAny, "A0", It.IsAny<string>())).Returns("FoundCategory");
            _mockSetting.Setup(m => m.GetDcCategoryVoltages("FoundCategory")).Returns(["HV", "LV"]);

            // Act
            _converter.WorkFlow(ref prodFlow);

            // Assert
            Assert.AreEqual("NV", row.Voltage);
        }

        [TestMethod]
        public void WorkFlow_SetsDefaultVoltage_WhenCategoryDoesNotExist_2()
        {
            // Arrange
            var row = new BistProdFlowRow
            {
                Pattern = "P1_ERT_2",
                Voltage = "NV",
                VoltageMode = "Nominal"
            };
            var prodFlow = new BistProdFlowSheet { Rows = [row] };
            prodFlow.MbistSheet.SheetName = "Module_A0";
            _mockSetting.Object.DcCategoryInfos = [];
            _mockSetting.Setup(m => m.FindMbistCatgeoryName(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), null, out It.Ref<EnumMessageLevel>.IsAny, out It.Ref<string>.IsAny, "A0", It.IsAny<string>())).Returns("FoundCategory");
            _mockSetting.Setup(m => m.GetDcCategoryVoltages("FoundCategory")).Returns(["HV"]);

            // Act
            _converter.WorkFlow(ref prodFlow);

            // Assert
            Assert.AreEqual("HV", row.Voltage);
        }

        private static BistProdFlowSheet CreateMockFlow(string pattern, string voltage)
        {
            var sheet = new BistProdFlowSheet
            {
                Rows =
                [
                    new() { Pattern = pattern, Voltage = voltage }
                ]
            };
            sheet.MbistSheet.SheetName = "Test_A1";
            return sheet;
        }
    }
}
