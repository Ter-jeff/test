using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class ChannelMapRowTests
    {
        [TestMethod]
        public void ChannelMapRow_DefaultConstructor_CreatesEmptyInstance()
        {
            // Arrange & Act
            var channelMapRow = new ChannelMapRow();

            // Assert
            Assert.IsInstanceOfType(channelMapRow, typeof(IgxlRow));
            Assert.IsNull(channelMapRow.ColumnA);
        }

        [TestMethod]
        public void ChannelMapRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var channelMapRow = new ChannelMapRow();

            // Assert
            Assert.IsInstanceOfType(channelMapRow, typeof(IgxlRow));
        }
    }
}
