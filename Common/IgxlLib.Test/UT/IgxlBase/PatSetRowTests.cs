using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class PatSetRowTests
    {
        [TestMethod]
        public void PatSetRow_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var patSetRow = new PatSetRow
            {
                // Act
                PatternSet = "PatSet1",
                TdGroup = "TDG1",
                TimeDomain = "TD1",
                Enable = "1",
                File = "pattern.pat",
                Burst = "BurstMode",
                StartLabel = "START",
                StopLabel = "STOP",
                Comment = "Pattern set comment"
            };

            // Assert
            Assert.AreEqual("PatSet1", patSetRow.PatternSet);
            Assert.AreEqual("TDG1", patSetRow.TdGroup);
            Assert.AreEqual("TD1", patSetRow.TimeDomain);
            Assert.AreEqual("1", patSetRow.Enable);
            Assert.AreEqual("pattern.pat", patSetRow.File);
            Assert.AreEqual("BurstMode", patSetRow.Burst);
            Assert.AreEqual("START", patSetRow.StartLabel);
            Assert.AreEqual("STOP", patSetRow.StopLabel);
        }

        [TestMethod]
        public void PatSetRow_CompareRow_WithIdenticalRows_ReturnsTrue()
        {
            // Arrange
            var row1 = new PatSetRow
            {
                TdGroup = "TDG1",
                TimeDomain = "TD1",
                Enable = "1",
                File = "pattern.pat",
                Burst = "BurstMode",
                Comment = "Test"
            };

            var row2 = new PatSetRow
            {
                TdGroup = "TDG1",
                TimeDomain = "TD1",
                Enable = "1",
                File = "pattern.pat",
                Burst = "BurstMode",
                Comment = "Test"
            };

            // Act
            bool result = row1.CompareRow(row2);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void PatSetRow_CompareRow_WithDifferentValues_ReturnsFalse()
        {
            // Arrange
            var row1 = new PatSetRow
            {
                TdGroup = "TDG1",
                TimeDomain = "TD1",
                Enable = "1",
                File = "pattern1.pat",
                Burst = "BurstMode",
                Comment = "Test"
            };

            var row2 = new PatSetRow
            {
                TdGroup = "TDG1",
                TimeDomain = "TD1",
                Enable = "1",
                File = "pattern2.pat",  // Different
                Burst = "BurstMode",
                Comment = "Test"
            };

            // Act
            bool result = row1.CompareRow(row2);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void PatSetRow_CompareRow_CaseInsensitive_ReturnsTrue()
        {
            // Arrange
            var row1 = new PatSetRow
            {
                TdGroup = "TDG1",
                TimeDomain = "td1",
                Enable = "yes",
                File = "pattern.pat",
                Burst = "burstmode",
                Comment = "test"
            };

            var row2 = new PatSetRow
            {
                TdGroup = "tdg1",
                TimeDomain = "TD1",
                Enable = "YES",
                File = "PATTERN.PAT",
                Burst = "BurstMode",
                Comment = "TEST"
            };

            // Act
            bool result = row1.CompareRow(row2);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void PatSetRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var patSetRow = new PatSetRow();

            // Assert
            Assert.IsInstanceOfType(patSetRow, typeof(IgxlRow));
        }
    }
}
