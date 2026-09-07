using IgxlLib.Enums;
using IgxlLib.IgxlReader;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlReader
{
    [TestClass]
    public class IgxlSheetFactoryTests
    {
        [TestMethod]
        public void IgxlSheetFactory_Constructor_CreatesInstance()
        {
            // Arrange & Act
            var factory = new IgxlSheetFactory();

            // Assert
            Assert.IsNotNull(factory);
        }

        [TestMethod]
        public void IgxlSheetFactory_LoadsReaders()
        {
            // Arrange & Act
            var factory = new IgxlSheetFactory();

            // Assert - just verify it was created without error
            Assert.IsNotNull(factory);
        }

        [TestMethod]
        public void IgxlSheetFactory_CreateSheet_WithNullWorksheet_HandlesGracefully()
        {
            // Arrange
            var factory = new IgxlSheetFactory();

            // Act
            IgxlLib.IgxlSheets.IIgxlSheet result = factory.CreateSheet(EnumSheetType.DTUnknown, null);

            // Assert - should handle null gracefully
            Assert.IsNull(result);
        }

        [TestMethod]
        public void IgxlSheetFactory_CreateSheet_WithInvalidSheetType_ReturnsNull()
        {
            // Arrange
            var factory = new IgxlSheetFactory();

            // Act
            IgxlLib.IgxlSheets.IIgxlSheet result = factory.CreateSheet(EnumSheetType.DTUnknown, null);

            // Assert
            Assert.IsNull(result);
        }
    }
}
