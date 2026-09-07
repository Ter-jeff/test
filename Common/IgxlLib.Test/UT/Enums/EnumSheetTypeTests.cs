using IgxlLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.Enums
{
    [TestClass]
    public class EnumSheetTypeTests
    {
        [TestMethod]
        public void EnumSheetType_CanCompareValues()
        {
            // Arrange
            EnumSheetType sheetType1 = EnumSheetType.DTFlowtableSheet;
            EnumSheetType sheetType2 = EnumSheetType.DTFlowtableSheet;
            EnumSheetType sheetType3 = EnumSheetType.DTLevelSheet;

            // Act & Assert
            Assert.AreEqual(sheetType1, sheetType2);
            Assert.AreNotEqual(sheetType1, sheetType3);
        }

        [TestMethod]
        public void EnumSheetType_CanCastToInt()
        {
            // Arrange & Act
            int value = (int)EnumSheetType.DTFlowtableSheet;

            // Assert
            Assert.IsTrue(value >= 0);
        }
    }
}
