using IgxlLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.Enums
{
    [TestClass]
    public class EnumJobTests
    {
        [TestMethod]
        public void EnumJob_CanCompareValues()
        {
            // Arrange
            EnumJob job1 = EnumJob.CP1;
            EnumJob job2 = EnumJob.CP1;
            EnumJob job3 = EnumJob.FT1;

            // Act & Assert
            Assert.AreEqual(job1, job2);
            Assert.AreNotEqual(job1, job3);
        }

        [TestMethod]
        public void EnumJob_HasNoneValue()
        {
            // Arrange & Act
            EnumJob noneJob = EnumJob.None;

            // Assert
            Assert.AreEqual(5, (int)noneJob);
        }

        [TestMethod]
        public void EnumJob_CanCastFromInt()
        {
            // Arrange & Act
            var job = (EnumJob)2;

            // Assert
            Assert.AreEqual(EnumJob.FT1, job);
        }
    }
}
