using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class MeasCPinCheckerTests : TestBase
    {
        private HardIpSheet _mockSheet = null!;
        private Mock<HardIpPattern> _mockPattern = null!;

        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            string hardIpInfoFile = Path.Combine(InputPath, "HardIp", "HardIP_PatInfo_Default.log");
            List<HardIpInfo> patInfo = new PatInfoReader().ExtractHardIpInfos(hardIpInfoFile);
            LocalSpecs.HardIpInfos = new HardIpInfos(patInfo);
        }

        [TestInitialize]
        public void Setup()
        {
            ErrorReportManager.ClearErrors();

            _mockSheet = new HardIpSheet
            {
                PlanHeaderIdx = new Dictionary<string, int> { { "measIndex", 5 } }
            };

            _mockPattern = new Mock<HardIpPattern>();
            _mockPattern.Object.Pattern = new PatternClass("cz_nvsa0_a_fulp_an_an00_mea_jtg_dio_allfrv_si_actcons_t7") { TestPlanPatternName = "cz_nvsa0_a_fulp_an_an00_mea_jtg_dio_allfrv_si_actcons_t7" };
            _mockPattern.Object.MeasPins = [];

            var mockPinMapSheet = new Mock<PinMapSheet>("SomeName");
            mockPinMapSheet.Setup(x => x.IsPinExist("jtag_tdo".ToUpper())).Returns(true);

            var kvp = new KeyValuePair<string, PinMapSheet>("TestKey", mockPinMapSheet.Object);
            TestProgram.IgxlWorkBk.Clear();
            TestProgram.IgxlWorkBk.PinMapPair = kvp;
        }

        [TestMethod]
        public void Check_Skips_WhenPatternIsInstance()
        {
            // Arrange
            _mockPattern.Object.Pattern.TestPlanPatternName = "Instance: MyPattern";
            var checker = new MeasCPinChecker(_mockSheet, _mockPattern.Object);

            // Act
            checker.Check();

            // Assert
            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_Skips_WhenPatternIsOpcode()
        {
            _mockPattern.Object.Pattern.TestPlanPatternName = "Opcode:TEST";
            var checker = new MeasCPinChecker(_mockSheet, _mockPattern.Object);

            checker.Check();

            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

        [TestMethod]
        public void Check_AddsError_NoMeasC()
        {
            // Arrange
            var mockPattern = new Mock<HardIpPattern>();
            mockPattern.Object.Pattern = new PatternClass("P1") { TestPlanPatternName = "P1" };

            var checker = new MeasCPinChecker(_mockSheet, mockPattern.Object);

            // Act
            checker.Check();

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.IsFalse(errors.Count != 0);
        }

        [TestMethod]
        public void Check_AddsError_WrongMeasPinInPatInfo()
        {
            // Arrange
            var measPin = new MeasPin { MeasType = "MeasC", PinName = "WrongPin", RowNum = 10 };
            _mockPattern.Object.MeasPins.Add(measPin);

            var checker = new MeasCPinChecker(_mockSheet, _mockPattern.Object);

            // Act
            checker.Check();

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.IsTrue(errors.Any(e => e.ErrorCode == HardIpErrorType.E_WrongMeasPinInPatInfo_01));
        }

        [TestMethod]
        public void Check_AddsError_MissingPatInfoPinInPinMap()
        {
            // Arrange
            var measPin = new MeasPin { MeasType = "MeasC", PinName = "WrongPin", RowNum = 10 };
            _mockPattern.Object.MeasPins.Add(measPin);
            var mockPinMapSheet = new Mock<PinMapSheet>("SomeName");
            mockPinMapSheet.Setup(x => x.IsPinExist(It.IsAny<string>())).Returns(false);

            var kvp = new KeyValuePair<string, PinMapSheet>("TestKey", mockPinMapSheet.Object);
            TestProgram.IgxlWorkBk.Clear();
            TestProgram.IgxlWorkBk.PinMapPair = kvp;

            var checker = new MeasCPinChecker(_mockSheet, _mockPattern.Object);

            // Act
            checker.Check();

            // Assert
            List<Error> errors = ErrorReportManager.GetErrorList();
            Assert.IsTrue(errors.Any(e => e.ErrorCode == HardIpErrorType.E_MissingPatInfoPinInPinMap_01));
        }

        [TestMethod]
        public void Check_AddsError_WhenInfoPinIsEmpty_AndNotMultiDomain()
        {
            var mockPattern = new Mock<HardIpPattern>();
            mockPattern.Object.Pattern =
                new PatternClass("normal_pattern_without_mt");

            mockPattern.Object.MeasPins =
            [
                new() { MeasType = "MeasC", PinName = "PIN1", RowNum = 5 }
            ];

            var checker = new MeasCPinChecker(_mockSheet, mockPattern.Object);

            checker.Check();

            Assert.IsTrue(ErrorReportManager.GetErrorList().Any(e => e.ErrorCode.FullCode == HardIpErrorType.E_WrongMeasC_01.FullCode));
        }

        [TestMethod]
        public void Check_NoError_WhenMeasCPinMatchesPatInfo()
        {
            var measPin = new MeasPin
            {
                MeasType = "MeasC",
                PinName = "jtag_tdo",
                RowNum = 10
            };

            _mockPattern.Object.MeasPins.Add(measPin);

            var checker = new MeasCPinChecker(_mockSheet, _mockPattern.Object);

            checker.Check();

            Assert.AreEqual(0, ErrorReportManager.GetErrorCount());
        }

    }
}
