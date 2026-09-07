using Automation.Utility;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Utility
{
    [TestClass]
    public class StepEvaluatorTests
    {
        [TestCleanup]
        public void Cleanup()
        {
            StepEvaluator.ErrMSg = "";
        }

        [TestMethod]
        public void Evaluate_SimpleNumericRange_ReturnsExpectedStepCount()
        {
            // Act
            int result = StepEvaluator.Evaluate("0,10,1");

            // Assert
            Assert.AreEqual(10, result);
        }

        [TestMethod]
        public void Evaluate_InvalidStartToken_SetsErrMsgAndReturnsZeroRange()
        {
            // Arrange - "abc" cannot be evaluated as a numeric expression by DataTable.Compute
            StepEvaluator.ErrMSg = "";

            // Act
            int result = StepEvaluator.Evaluate("abc,10,1");

            // Assert
            Assert.AreEqual(0, result);
            Assert.IsTrue(StepEvaluator.ErrMSg.Contains("Error!"));
        }

        [TestMethod]
        public void CalStr_PlainNumber_ReturnsUnchanged()
        {
            // Act
            string result = StepEvaluator.CalStr(["100"]);

            // Assert
            Assert.AreEqual("100", result);
        }

        [TestMethod]
        public void CalStr_WithMilliUnitSuffix_SubstitutesScaleFactor()
        {
            // Act - the "V" unit is stripped first, leaving "10m" which matches the
            // digit+scale-letter pattern and gets expanded to a multiplication by 10^-3
            string result = StepEvaluator.CalStr(["10mV"]);

            // Assert
            Assert.AreEqual("10*0.001", result);
        }

        [TestMethod]
        public void CalStr_MultipleItems_ConcatenatesResults()
        {
            // Act
            string result = StepEvaluator.CalStr(["1", "2"]);

            // Assert
            Assert.AreEqual("12", result);
        }

        [TestMethod]
        public void Replace_MinusSign_InsertsLeadingComma()
        {
            // Act
            string result = StepEvaluator.Replace("1-2", "-");

            // Assert
            Assert.AreEqual("1,-2", result);
        }

        [TestMethod]
        public void Replace_PlusSign_InsertsLeadingComma()
        {
            // Act
            string result = StepEvaluator.Replace("1+2", "+");

            // Assert
            Assert.AreEqual("1,+2", result);
        }

        [TestMethod]
        public void GetCalStr_BuildsSubtractionExpressionFromBothLists()
        {
            // Act
            string result = StepEvaluator.GetCalStr(["10"], ["3"]);

            // Assert
            Assert.AreEqual("3-(10)", result);
        }
    }
}
