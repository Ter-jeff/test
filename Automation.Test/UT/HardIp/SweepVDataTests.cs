using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class SweepVDataTests : FunctionTestBase
    {
        [TestMethod]
        public void Operand_ShouldBePlus_WhenStartLessThanStop()
        {
            var data = new SweepVData("PIN,1,2,0.1");
            Assert.AreEqual("+", data.Operand);
            Assert.AreEqual("<", data.Comparator);
            Assert.IsTrue(data.CheckStep);
        }

        [TestMethod]
        public void Operand_ShouldBeMinus_WhenStartGreaterThanStop()
        {
            var data = new SweepVData("PIN,2,1,-0.1");
            Assert.AreEqual("-", data.Operand);
            Assert.AreEqual(">", data.Comparator);
            Assert.IsTrue(data.CheckStep);
        }

        [TestMethod]
        public void CheckStep_ShouldBeFalse_WhenDirectionMismatch()
        {
            var data = new SweepVData("PIN,1,2,-0.1");
            Assert.IsFalse(data.CheckStep);
        }

        [TestMethod]
        public void Constructor_ShouldParseSweepStrCorrectly()
        {
            var sweep = new SweepVData("PIN:1,2,0.5", "LabelX");
            Assert.AreEqual("PIN", sweep.PinName);
            Assert.AreEqual("LabelX", sweep.Axis);
            Assert.AreEqual("1", sweep.Start);
            Assert.AreEqual("2", sweep.Stop);
            Assert.AreEqual("0.5", sweep.Step);
        }

        [TestMethod]
        public void Operand_ShouldSetIsEquation_WhenNonNumeric()
        {
            var data = new SweepVData("PIN,X,Y,0.1");
            _ = data.Operand;
            Assert.IsTrue(data.IsEquation);
        }

        [TestMethod]
        public void Constructor_WithVoltageList_ShouldSetFields()
        {
            var data = new SweepVData("VDD", "0.9,1.0,1.1", "LabelV", "1", "VDD_MAIN", "FORCE_TYPE");
            Assert.AreEqual("VDD", data.PinName);
            Assert.AreEqual("0.9,1.0,1.1", data.VoltageList);
            Assert.AreEqual("LabelV", data.Axis);
            Assert.AreEqual("1", data.Order);
            Assert.AreEqual("VDD_MAIN", data.InstanceVoltage);
            Assert.AreEqual("FORCE_TYPE", data.ForceType);
            Assert.AreEqual("LabelV", data.Type);
        }

        [TestMethod]
        public void SweepVData_Constructor_ShouldParseStandardSweepString()
        {
            // Arrange
            string sweepStr = "VDD_CORE:A+0.8,1.2,0.1";
            string label = "X";

            // Act
            var result = new SweepVData(sweepStr, label);

            // Assert
            Assert.AreEqual("X", result.Axis);
            Assert.AreEqual("", result.Type, "Type should be empty when IsDCSpecSweep is true.");
            Assert.AreNotEqual(null, result.PinName);
            Assert.AreEqual("A+0.8", result.Start);
            Assert.AreEqual("1.2", result.Stop);
        }

        [TestMethod]
        public void Constructor_ShouldInitialize_WithLoopStr()
        {
            SweepVData data = new SweepVData("1,2,3");

            Assert.AreNotEqual(null, data);
        }

        [TestMethod]
        public void Operand_ShouldSetIsEquation_WhenInvalidNumber()
        {
            SweepVData data = new SweepVData("PIN,A,2,1");

            string result = data.Operand;

            Assert.AreEqual("+", result);
            Assert.IsTrue(data.IsEquation);
        }

        [TestMethod]
        public void Comparator_ShouldReturnGreater_WhenStartGreaterThanStop()
        {
            SweepVData data = new SweepVData("PIN,3,1,1");

            string result = data.Comparator;

            Assert.AreEqual(">", result);
        }

        [TestMethod]
        public void CheckStep_ShouldReturnFalse_WhenOperandMinus_ButStepPositive()
        {
            SweepVData data = new SweepVData("PIN,3,1,1");

            bool result = data.CheckStep;

            Assert.IsFalse(result);
        }

        [TestMethod]
        public void CheckStep_ShouldReturnTrue_WhenValid()
        {
            SweepVData data = new SweepVData("PIN,1,2,1");

            bool result = data.CheckStep;

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void Constructor_ShouldParsePinStartStopStep()
        {
            SweepVData data = new SweepVData("PIN,1,2,0.5");

            Assert.AreEqual("PIN", data.PinName);
            Assert.AreEqual("1", data.Start);
            Assert.AreEqual("2", data.Stop);
            Assert.AreEqual("0.5", data.Step);
        }

        [TestMethod]
        public void Constructor2_ShouldParseSweepString()
        {
            SweepVData data = new SweepVData("PIN:1,2,1", "X");

            Assert.AreEqual("PIN", data.PinName);
            Assert.AreEqual("1", data.Start);
            Assert.AreEqual("2", data.Stop);
            Assert.AreEqual("1", data.Step);
        }

        [TestMethod]
        public void Constructor2_ShouldSetZero_WhenInvalidFormat()
        {
            SweepVData data = new SweepVData("PIN:1,2", "X");

            Assert.AreEqual("0", data.Start);
            Assert.AreEqual("0", data.Stop);
            Assert.AreEqual("0", data.Step);
        }

        [TestMethod]
        public void CopyConstructor_ShouldCloneAllFields()
        {
            SweepVData src = new SweepVData("PIN,1,2,1")
            {
                Type = "TEST",
                Order = "A"
            };

            SweepVData copy = new SweepVData(src);

            Assert.AreEqual(src.PinName, copy.PinName);
            Assert.AreEqual(src.Type, copy.Type);
            Assert.AreEqual(src.Start, copy.Start);
            Assert.AreEqual(src.Stop, copy.Stop);
            Assert.AreEqual(src.Step, copy.Step);
            Assert.AreEqual(src.Order, copy.Order);
        }

        [TestMethod]
        public void Copy_ShouldReturnNewInstance()
        {
            SweepVData src = new SweepVData("PIN,1,2,1");

            SweepVData copy = src.Copy();

            Assert.AreNotSame(src, copy);
            Assert.AreEqual(src.Start, copy.Start);
        }

        [TestMethod]
        public void CopyConstructor_ShouldHandleNull()
        {
            SweepVData copy = new SweepVData(null);

            Assert.AreEqual(null, copy.PinName);
        }

        [TestMethod]
        public void CheckStep_ShouldHandleNullStep()
        {
            SweepVData data = new SweepVData("PIN,1,2,1")
            {
                Step = null
            };

            bool result = data.CheckStep;

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void Comparator_ShouldFallback_WhenInvalidNumber()
        {
            SweepVData data = new SweepVData("PIN,A,B,1");

            string result = data.Comparator;

            Assert.AreEqual("<", result);
        }

        [TestMethod]
        public void Constructor2_ShouldSetIsEquation_WhenStartInvalidStopValid()
        {
            SweepVData data = new SweepVData("PIN:A,2,1", "X");

            Assert.IsTrue(data.IsEquation);
        }

        [TestMethod]
        public void Constructor2_ShouldSetType_WhenNotDcSpecSweep()
        {
            SweepVData data = new SweepVData("PIN:1,2,1", "X", false);

            Assert.AreEqual("SrcCodeIndexX", data.Type);
        }

    }
}
