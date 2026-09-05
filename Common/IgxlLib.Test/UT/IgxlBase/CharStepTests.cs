using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class CharStepTests
    {
        [TestMethod]
        public void CharStep_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var charStep = new CharStep("", "");

            // Assert
            Assert.AreEqual("", charStep.VoltageType);
            Assert.AreEqual("", charStep.StepName);
            Assert.AreEqual("", charStep.Mode);
        }

        [TestMethod]
        public void CharStep_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var charStep = new CharStep("Setup1", "Step1")
            {
                // Act
                VoltageType = "DC",
                Mode = "SweepMode",
                ParameterName = "Voltage",
                RangeFrom = "0",
                RangeTo = "5"
            };

            // Assert
            Assert.AreEqual("DC", charStep.VoltageType);
            Assert.AreEqual("Setup1", charStep.SetupName);
            Assert.AreEqual("Step1", charStep.StepName);
            Assert.AreEqual("SweepMode", charStep.Mode);
            Assert.AreEqual("Voltage", charStep.ParameterName);
            Assert.AreEqual("0", charStep.RangeFrom);
            Assert.AreEqual("5", charStep.RangeTo);
        }

        [TestMethod]
        public void CharStep_SetTestLimits_UpdatesLimitProperties()
        {
            // Arrange
            var charStep = new CharStep("", "")
            {
                // Act
                TestLimitLow = "1.0",
                TestLimitHigh = "10.0"
            };

            // Assert
            Assert.AreEqual("1.0", charStep.TestLimitLow);
            Assert.AreEqual("10.0", charStep.TestLimitHigh);
        }

        [TestMethod]
        public void CharStep_SetAlgorithmProperties_UpdatesCorrectly()
        {
            // Arrange
            var charStep = new CharStep("", "")
            {
                // Act
                AlgorithmName = "Binary Search",
                AlgorithmArguments = "Arg1,Arg2",
                AlgorithmResultsCheck = "Check1"
            };

            // Assert
            Assert.AreEqual("Binary Search", charStep.AlgorithmName);
            Assert.AreEqual("Arg1,Arg2", charStep.AlgorithmArguments);
            Assert.AreEqual("Check1", charStep.AlgorithmResultsCheck);
        }

        [TestMethod]
        public void CharStep_SetApplyToProperties_UpdatesCorrectly()
        {
            // Arrange
            var charStep = new CharStep("", "")
            {
                // Act
                ApplyToPinExecMode = "Parallel",
                ApplyToPins = "Pin1,Pin2,Pin3",
                ApplyToTimeSets = "TimeSet1"
            };

            // Assert
            Assert.AreEqual("Parallel", charStep.ApplyToPinExecMode);
            Assert.AreEqual("Pin1,Pin2,Pin3", charStep.ApplyToPins);
            Assert.AreEqual("TimeSet1", charStep.ApplyToTimeSets);
        }

        [TestMethod]
        public void CharStep_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var charStep = new CharStep("", "");

            // Assert
            Assert.IsInstanceOfType(charStep, typeof(IgxlRow));
        }
    }
}
