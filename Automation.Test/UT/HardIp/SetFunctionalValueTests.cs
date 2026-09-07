using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;

using CommonReaderLib.PatternListCsv;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SetFunctionalValueTests
    {
        private SetFunctionalValue _setValue = null!;

        [TestInitialize]
        public void Setup()
        {
            HardIpParaData paraData = new(EnumBlock.HardIp);
            var dummySheet = new HardIpSheet();
            var inputData = new HardIpInputData(paraData);
            _setValue = new SetFunctionalValue(inputData, dummySheet);
        }

        private static (string Pin, string Equation, string Assignment) InvokeGetDigSrc(SetFunctionalValue setValue, HardIpPattern pattern)
        {
            string pin = "";
            string equation = "";
            string assignment = "";
            setValue.GetDigSrc(pattern, ref pin, ref equation, ref assignment);
            return (pin, equation, assignment);
        }

        private static HardIpPattern NewPattern(string realPatternName)
        {
            return new HardIpPattern { SheetName = "Sheet1", Pattern = new PatternClass(realPatternName) };
        }

        [TestMethod]
        public void GetDigSrc_NoMatchingSelsramPattern_ReturnsAllEmpty()
        {
            // Arrange
            HardIpPattern pattern = NewPattern("TEST_NORMAL");

            // Act
            (string pin, string equation, string assignment) = InvokeGetDigSrc(_setValue, pattern);

            // Assert
            Assert.AreEqual("", pin);
            Assert.AreEqual("", equation);
            Assert.AreEqual("", assignment);
        }

        [TestMethod]
        public void GetDigSrc_SingleMatchingSelsramPattern_SetsPinEquationAndAssignment()
        {
            // Arrange
            HardIpPattern pattern = NewPattern("TEST_SRAMDSSC");

            // Act
            (string pin, string equation, string assignment) = InvokeGetDigSrc(_setValue, pattern);

            // Assert
            Assert.AreEqual("JTAG_TDI", pin);
            Assert.AreEqual("C", equation);
            Assert.AreEqual("C=Selsram()", assignment);
        }

        [TestMethod]
        public void GetDigSrc_MixOfMatchingAndNonMatchingPatterns_BuildsPipeJoinedEquation()
        {
            // Arrange - two burst-pattern segments, only the second matches the selsram regex
            HardIpPattern pattern = NewPattern("TEST_NORMAL+TEST_SRAMDSSC");

            // Act
            (string pin, string equation, string assignment) = InvokeGetDigSrc(_setValue, pattern);

            // Assert
            Assert.AreEqual("JTAG_TDI", pin);
            Assert.AreEqual("|C", equation);
            Assert.AreEqual("C=Selsram()", assignment);
        }

        [TestMethod]
        public void GetDigSrc_MultipleMatchingPatterns_DedupesAssignmentButKeepsEachEquationEntry()
        {
            // Arrange - two distinct burst-pattern groups both matching, the assignment set dedupes
            HardIpPattern pattern = NewPattern("TEST_SRAMDSSC,OTHER_SRAMDSSC");

            // Act
            (string pin, string equation, string assignment) = InvokeGetDigSrc(_setValue, pattern);

            // Assert
            Assert.AreEqual("JTAG_TDI", pin);
            Assert.AreEqual("C|C", equation);
            Assert.AreEqual("C=Selsram()", assignment);
        }

        [TestMethod]
        public void GetDigSrc_EmptyPatternName_ReturnsAllEmpty()
        {
            // Arrange
            HardIpPattern pattern = NewPattern("");

            // Act
            (string pin, string equation, string assignment) = InvokeGetDigSrc(_setValue, pattern);

            // Assert
            Assert.AreEqual("", pin);
            Assert.AreEqual("", equation);
            Assert.AreEqual("", assignment);
        }
    }
}
