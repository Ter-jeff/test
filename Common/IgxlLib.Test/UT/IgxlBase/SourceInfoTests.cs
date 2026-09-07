using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class SourceInfoTests
    {
        [TestMethod]
        public void Constructor_Default_InitializesWithEmptyStrings()
        {
            // Act
            var sourceInfo = new SourceInfo();

            // Assert
            Assert.IsNotNull(sourceInfo);
            Assert.AreEqual(string.Empty, sourceInfo.Name);
            Assert.AreEqual(string.Empty, sourceInfo.Block);
        }

        [TestMethod]
        public void Constructor_Copy_WithValidInstance_CopiesAllProperties()
        {
            // Arrange
            var original = new SourceInfo
            {
                Name = "SignalA",
                Block = "SocBlock"
            };

            // Act
            var copy = new SourceInfo(original);

            // Assert
            Assert.IsNotNull(copy);
            Assert.AreEqual(original.Name, copy.Name);
            Assert.AreEqual(original.Block, copy.Block);
            Assert.AreNotSame(original, copy);
        }

        [TestMethod]
        public void Constructor_Copy_WithNullInstance_GracefullyHandlesAndKeepsDefaults()
        {
            // Act
            var copy = new SourceInfo(null);

            // Assert
            Assert.IsNotNull(copy);
            Assert.AreEqual(string.Empty, copy.Name);
            Assert.AreEqual(string.Empty, copy.Block);
        }

        [TestMethod]
        public void Copy_Method_ReturnsNewIdenticalInstance()
        {
            // Arrange
            var original = new SourceInfo
            {
                Name = "SignalB",
                Block = "CpuBlock"
            };

            // Act
            SourceInfo copy = original.Copy();

            // Assert
            Assert.IsNotNull(copy);
            Assert.AreEqual(original.Name, copy.Name);
            Assert.AreEqual(original.Block, copy.Block);
            Assert.AreNotSame(original, copy);
        }
    }
}
