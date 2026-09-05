using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{

    [TestClass]
    public class TimingRowTests
    {
        [TestMethod]
        public void TimingRow_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var timingRow = new TimingRow
            {
                // Act
                PinGrpName = "Data_Group",
                PinGrpClockPeriod = "10ns",
                PinGrpSetup = "Setup1",
                DataSrc = "PatternFile",
                DataFmt = "Binary",
                DriveOn = "0ns",
                DriveData = "5ns",
                CompareMode = "Match",
                EdgeMode = "Rising",
                Comment = "Test timing"
            };

            // Assert
            Assert.AreEqual("Data_Group", timingRow.PinGrpName);
            Assert.AreEqual("10ns", timingRow.PinGrpClockPeriod);
            Assert.AreEqual("Setup1", timingRow.PinGrpSetup);
            Assert.AreEqual("PatternFile", timingRow.DataSrc);
            Assert.AreEqual("Binary", timingRow.DataFmt);
            Assert.AreEqual("0ns", timingRow.DriveOn);
            Assert.AreEqual("5ns", timingRow.DriveData);
            Assert.AreEqual("Match", timingRow.CompareMode);
            Assert.AreEqual("Rising", timingRow.EdgeMode);
            Assert.AreEqual("Test timing", timingRow.Comment);
        }

        [TestMethod]
        public void TimingRow_CopyConstructor_CopiesAllProperties()
        {
            // Arrange
            var originalRow = new TimingRow
            {
                PinGrpName = "Clock_Group",
                PinGrpClockPeriod = "20ns",
                PinGrpSetup = "Setup2",
                DataSrc = "Vector",
                DataFmt = "Hex",
                DriveOn = "2ns",
                DriveData = "10ns",
                DriveReturn = "15ns",
                DriveOff = "18ns",
                CompareMode = "Verify",
                CompareOpen = "1ns",
                CompareClose = "12ns",
                EdgeMode = "Falling",
                Comment = "Clock timing",
            };

            // Act
            var copiedRow = new TimingRow(originalRow);

            // Assert
            Assert.AreEqual(originalRow.PinGrpName, copiedRow.PinGrpName);
            Assert.AreEqual(originalRow.PinGrpClockPeriod, copiedRow.PinGrpClockPeriod);
            Assert.AreEqual(originalRow.DataSrc, copiedRow.DataSrc);
            Assert.AreEqual(originalRow.DriveOn, copiedRow.DriveOn);
            Assert.AreEqual(originalRow.Comment, copiedRow.Comment);
            Assert.AreEqual(originalRow.Name, copiedRow.Name);
            Assert.AreEqual(originalRow.CyclePeriod, copiedRow.CyclePeriod);
        }

        [TestMethod]
        public void TimingRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var timingRow = new TimingRow();

            // Assert
            Assert.IsInstanceOfType(timingRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void TimingRow_MultipleInstances_AreIndependent()
        {
            // Arrange & Act
            var row1 = new TimingRow { PinGrpName = "Group1", PinGrpClockPeriod = "10ns" };
            var row2 = new TimingRow { PinGrpName = "Group2", PinGrpClockPeriod = "20ns" };

            // Assert
            Assert.AreEqual("Group1", row1.PinGrpName);
            Assert.AreEqual("Group2", row2.PinGrpName);
            Assert.AreNotEqual(row1.PinGrpClockPeriod, row2.PinGrpClockPeriod);
        }
    }
}
