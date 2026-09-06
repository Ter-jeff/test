using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class IgxlRowTests
    {
        [TestMethod]
        public void IgxlRow_Constructor_InitializesPropertiesCorrectly()
        {
            // Arrange & Act
            var row = new IgxlRow();

            // Assert
            Assert.IsNull(row.ColumnA);
            Assert.IsNull(row.SheetName);
            Assert.AreEqual(0, row.RowNum);
            Assert.IsFalse(row.IsBackup);
        }

        [TestMethod]
        public void IgxlRow_SetProperties_UpdatesPropertiesCorrectly()
        {
            // Arrange
            var row = new IgxlRow();
            string columnAValue = "TestColumn";
            string sheetName = "TestSheet";
            int rowNum = 5;
            bool isBackup = true;

            // Act
            row.ColumnA = columnAValue;
            row.SheetName = sheetName;
            row.RowNum = rowNum;
            row.IsBackup = isBackup;

            // Assert
            Assert.AreEqual(columnAValue, row.ColumnA);
            Assert.AreEqual(sheetName, row.SheetName);
            Assert.AreEqual(rowNum, row.RowNum);
            Assert.IsTrue(row.IsBackup);
        }

        [TestMethod]
        public void IgxlRow_MultipleInstances_AreIndependent()
        {
            // Arrange
            var row1 = new IgxlRow { ColumnA = "A", SheetName = "Sheet1" };
            var row2 = new IgxlRow { ColumnA = "B", SheetName = "Sheet2" };

            // Assert
            Assert.AreEqual("A", row1.ColumnA);
            Assert.AreEqual("Sheet1", row1.SheetName);
            Assert.AreEqual("B", row2.ColumnA);
            Assert.AreEqual("Sheet2", row2.SheetName);
        }
    }
}
