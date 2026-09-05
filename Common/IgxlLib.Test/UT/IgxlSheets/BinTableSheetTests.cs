using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class BinTableSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void BinTableSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "BinTable";

            // Act
            var binTableSheet = new BinTableSheet(sheetName);

            // Assert
            Assert.IsNotNull(binTableSheet);
            Assert.AreEqual(sheetName, binTableSheet.Name);
            Assert.AreEqual("DTBintablesSheet", binTableSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.BinTable, binTableSheet.IgxlSheetName);
            Assert.AreEqual(0, binTableSheet.Rows.Count);
        }

        [TestMethod]
        public void BinTableSheet_AddRow()
        {
            // Arrange
            var binTableSheet = new BinTableSheet("BinTable");
            var binRow = new BinTableRow
            {
                Name = "Bin1",
                ItemList = "Item0",
                Op = "GT",
                Sort = "1",
                Bin = "100",
                Result = "Pass",
                Comment = "Good test"
            };

            // Act
            binTableSheet.AddRow(binRow);

            // Assert
            Assert.AreEqual(1, binTableSheet.Rows.Count);
            Assert.AreEqual("Bin1", binTableSheet.Rows[0].Name);
        }

        [TestMethod]
        public void BinTableSheet_AddRows()
        {
            // Arrange
            var binTableSheet = new BinTableSheet("BinTable");
            var rows = new List<BinTableRow>
            {
                new() { Name = "Bin1", ItemList = "Item0", Op = "GT", Sort = "1", Bin = "100" },
                new() { Name = "Bin2", ItemList = "Item1", Op = "LT", Sort = "2", Bin = "200" },
                new() { Name = "Bin3", ItemList = "Item2", Op = "EQ", Sort = "3", Bin = "300" }
            };

            // Act
            binTableSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, binTableSheet.Rows.Count);
        }

        [TestMethod]
        public void BinTableSheet_RemoveRow()
        {
            // Arrange
            var binTableSheet = new BinTableSheet("BinTable");
            var row1 = new BinTableRow { Name = "Bin1", ItemList = "Item0" };
            var row2 = new BinTableRow { Name = "Bin2", ItemList = "Item1" };
            binTableSheet.AddRow(row1);
            binTableSheet.AddRow(row2);

            // Act
            binTableSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, binTableSheet.Rows.Count);
            Assert.AreEqual("Bin2", binTableSheet.Rows[0].Name);
        }

        [TestMethod]
        public void BinTableSheet_InsertRow()
        {
            // Arrange
            var binTableSheet = new BinTableSheet("BinTable");
            var row1 = new BinTableRow { Name = "Bin1", ItemList = "Item0" };
            var row2 = new BinTableRow { Name = "Bin2", ItemList = "Item1" };
            var rowToInsert = new BinTableRow { Name = "Bin3", ItemList = "Item2" };
            binTableSheet.AddRow(row1);
            binTableSheet.AddRow(row2);

            // Act
            int index = binTableSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, binTableSheet.Rows.Count);
            Assert.AreEqual("Bin3", binTableSheet.Rows[1].Name);
        }

        [TestMethod]
        public void BinTableSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var binTableSheet = new BinTableSheet("BinTable");

            // Assert
            Assert.IsNotNull(binTableSheet.GetErrors());
            Assert.AreEqual(0, binTableSheet.GetErrors().Count);
        }

        [TestMethod]
        public void BinTableSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var binTableSheet = new BinTableSheet("BinTable");

            // Assert
            Assert.AreEqual("DTBintablesSheet", binTableSheet.SheetType);
        }

        [TestMethod]
        public void BinTableSheet_Name_CanBeSet()
        {
            // Arrange
            var binTableSheet = new BinTableSheet("BinTable")
            {
                // Act
                Name = "NewBinTableName"
            };

            // Assert
            Assert.AreEqual("NewBinTableName", binTableSheet.Name);
        }

        [TestMethod]
        public void BinTableSheet_MultipleRows_WithDifferentOperators()
        {
            // Arrange
            var binTableSheet = new BinTableSheet("BinTable");
            string[] operators = ["GT", "LT", "EQ", "GE", "LE", "NE"];

            // Act
            for (int i = 0; i < operators.Length; i++)
            {
                var row = new BinTableRow
                {
                    Name = $"Bin{i}",
                    ItemList = $"Item{i}",
                    Op = operators[i],
                    Sort = (i + 1).ToString(),
                    Bin = ((i + 1) * 100).ToString()
                };
                binTableSheet.AddRow(row);
            }

            // Assert
            Assert.AreEqual(operators.Length, binTableSheet.Rows.Count);
            for (int i = 0; i < operators.Length; i++)
            {
                Assert.AreEqual(operators[i], binTableSheet.Rows[i].Op);
            }
        }
    }
}
