using IgxlLib.IgxlBase;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IgxlLib.Test.UT.IgxlBase
{
    [TestClass]
    public class BinTableRowTests
    {
        [TestMethod]
        public void BinTableRow_DefaultConstructor_InitializesCollections()
        {
            // Arrange & Act
            var binTableRow = new BinTableRow();

            // Assert
            Assert.AreEqual(0, binTableRow.ExtraBinDictionary.Count);
            Assert.AreEqual(0, binTableRow.Items.Count);
            Assert.AreEqual("", binTableRow.Name);
        }

        [TestMethod]
        public void BinTableRow_Constructor_WithSheetName_InitializesProperties()
        {
            // Arrange & Act
            var binTableRow = new BinTableRow("BinSheet1");

            // Assert
            Assert.AreEqual("BinSheet1", binTableRow.SheetName);
            Assert.AreEqual(0, binTableRow.RowNum);
        }

        [TestMethod]
        public void BinTableRow_SetProperties_UpdatesValuesCorrectly()
        {
            // Arrange
            var binTableRow = new BinTableRow
            {
                // Act
                Name = "BinRow1",
                ItemList = "Item1,Item2,Item3",
                Op = "AND",
                Sort = "Sort1",
                Bin = "100",
                Result = "Pass",
                Comment = "Bin comment"
            };

            // Assert
            Assert.AreEqual("BinRow1", binTableRow.Name);
            Assert.AreEqual("Item1,Item2,Item3", binTableRow.ItemList);
            Assert.AreEqual("AND", binTableRow.Op);
            Assert.AreEqual("Sort1", binTableRow.Sort);
            Assert.AreEqual("100", binTableRow.Bin);
            Assert.AreEqual("Pass", binTableRow.Result);
        }

        [TestMethod]
        public void BinTableRow_AddItemsToList_UpdatesList()
        {
            // Arrange
            var binTableRow = new BinTableRow();

            // Act
            binTableRow.Items.Add("Item1");
            binTableRow.Items.Add("Item2");
            binTableRow.ItemsWithIndex[0] = "Item1";
            binTableRow.ItemsWithIndex[1] = "Item2";

            // Assert
            Assert.AreEqual(2, binTableRow.Items.Count);
            Assert.AreEqual(2, binTableRow.ItemsWithIndex.Count);
        }

        [TestMethod]
        public void BinTableRow_AddExtraBinValues_UpdatesDictionary()
        {
            // Arrange
            var binTableRow = new BinTableRow();

            // Act
            binTableRow.ExtraBinDictionary["Extra1"] = "Value1";
            binTableRow.ExtraBinDictionary["Extra2"] = "Value2";

            // Assert
            Assert.AreEqual(2, binTableRow.ExtraBinDictionary.Count);
            Assert.AreEqual("Value1", binTableRow.ExtraBinDictionary["Extra1"]);
        }

        [TestMethod]
        public void BinTableRow_Inherits_FromIgxlRow()
        {
            // Arrange & Act
            var binTableRow = new BinTableRow();

            // Assert
            Assert.IsInstanceOfType(binTableRow, typeof(IgxlRow));
        }

        [TestMethod]
        public void BinTableRow_CopyConstructor_CopiesAllProperties()
        {
            // Arrange
            var original = new BinTableRow("Sheet1")
            {
                Name = "BinRow1",
                ItemList = "Item1,Item2",
                Op = "AND",
                Sort = "Sort1",
                Bin = "100",
                Result = "Pass",
                Comment = "Test"
            };
            original.Items.Add("Item1");
            original.Items.Add("Item2");
            original.ExtraBinDictionary["Extra"] = "Value";

            // Act
            var copy = new BinTableRow(original);

            // Assert
            Assert.AreEqual(original.Name, copy.Name);
            Assert.AreEqual(original.ItemList, copy.ItemList);
            Assert.AreEqual(original.Op, copy.Op);
            Assert.AreEqual(original.Sort, copy.Sort);
            Assert.AreEqual(original.Bin, copy.Bin);
            Assert.AreEqual(original.Result, copy.Result);
            Assert.AreEqual(2, copy.Items.Count);
            Assert.AreEqual(1, copy.ExtraBinDictionary.Count);
        }

        [TestMethod]
        public void BinTableRow_Copy_Method_CreatesIndependentCopy()
        {
            // Arrange
            var original = new BinTableRow("Sheet1")
            {
                Name = "BinRow1",
                Bin = "100"
            };
            original.Items.Add("Item1");

            // Act
            BinTableRow copy = original.Copy();

            // Assert
            Assert.AreEqual(original.Name, copy.Name);
            Assert.AreEqual(original.Bin, copy.Bin);
            Assert.AreEqual(1, copy.Items.Count);
            // Verify they are different objects
            Assert.AreNotSame(original, copy);
            Assert.AreNotSame(original.Items, copy.Items);
        }

        [TestMethod]
        public void BinTableRow_GetUniqKey_WithItems()
        {
            // Arrange
            var binTableRow = new BinTableRow
            {
                Name = "BinRow1",
                ItemList = "ItemList1"
            };
            binTableRow.Items.Add("Item1");
            binTableRow.Items.Add("Item2");

            // Act
            string key = binTableRow.GetUniqKey();

            // Assert
            Assert.IsNotNull(key);
            Assert.IsTrue(key.Contains("binrow1"));
            Assert.IsTrue(key.Contains("itemlist1"));
            Assert.IsTrue(key.Contains("item1"));
            Assert.IsTrue(key.Contains("item2"));
        }

        [TestMethod]
        public void BinTableRow_GetUniqKey_EmptyItems()
        {
            // Arrange
            var binTableRow = new BinTableRow
            {
                Name = "BinRow1",
                ItemList = "ItemList1"
            };

            // Act
            string key = binTableRow.GetUniqKey();

            // Assert
            Assert.IsNotNull(key);
            Assert.IsTrue(key.Contains("binrow1"));
        }

        [TestMethod]
        public void BinTableRow_AllProperties_CanBeSet()
        {
            // Arrange
            var binTableRow = new BinTableRow
            {
                // Act
                Name = "Test",
                ItemList = "Items",
                Op = "OR",
                Sort = "SortKey",
                Bin = "200",
                Result = "Fail",
                Comment = "Comment",
                SheetName = "Sheet",
                RowNum = 5
            };

            // Assert
            Assert.AreEqual("Test", binTableRow.Name);
            Assert.AreEqual("Items", binTableRow.ItemList);
            Assert.AreEqual("OR", binTableRow.Op);
            Assert.AreEqual("SortKey", binTableRow.Sort);
            Assert.AreEqual("200", binTableRow.Bin);
            Assert.AreEqual("Fail", binTableRow.Result);
            Assert.AreEqual("Comment", binTableRow.Comment);
            Assert.AreEqual("Sheet", binTableRow.SheetName);
            Assert.AreEqual(5, binTableRow.RowNum);
        }

        [TestMethod]
        public void BinTableRow_Multiple_Items_Operations()
        {
            // Arrange
            var binTableRow = new BinTableRow();

            // Act
            for (int i = 0; i < 10; i++)
            {
                binTableRow.Items.Add($"Item{i}");
                binTableRow.ItemsWithIndex[i] = $"Item{i}";
            }

            // Assert
            Assert.AreEqual(10, binTableRow.Items.Count);
            Assert.AreEqual(10, binTableRow.ItemsWithIndex.Count);
            Assert.AreEqual("Item5", binTableRow.Items[5]);
            Assert.AreEqual("Item5", binTableRow.ItemsWithIndex[5]);
        }

        [TestMethod]
        public void BinTableRow_CopyConstructor_WithNullSource()
        {
            // Arrange & Act
            var copy = new BinTableRow("");

            // Assert
            Assert.IsNotNull(copy);
        }

        [TestMethod]
        public void BinTableRow_ExtraBinDictionary_MultipleValues()
        {
            // Arrange
            var binTableRow = new BinTableRow();

            // Act
            for (int i = 0; i < 5; i++)
            {
                binTableRow.ExtraBinDictionary[$"Key{i}"] = $"Value{i}";
            }

            // Assert
            Assert.AreEqual(5, binTableRow.ExtraBinDictionary.Count);
            Assert.AreEqual("Value2", binTableRow.ExtraBinDictionary["Key2"]);
        }

        [TestMethod]
        public void BinTableRow_GetUniqKey_CaseInsensitive()
        {
            // Arrange
            var binTableRow1 = new BinTableRow { Name = "BINROW1", ItemList = "ITEMLIST1" };
            binTableRow1.Items.Add("ITEM1");
            var binTableRow2 = new BinTableRow { Name = "binrow1", ItemList = "itemlist1" };
            binTableRow2.Items.Add("item1");

            // Act
            string key1 = binTableRow1.GetUniqKey();
            string key2 = binTableRow2.GetUniqKey();

            // Assert
            Assert.AreEqual(key1, key2);
        }

        [TestMethod]
        public void BinTableRow_ResultFailStop_Constant()
        {
            // Arrange & Act
            string failStop = BinTableRow.ResultFailStop;

            // Assert
            Assert.AreEqual("Fail-stop", failStop);
        }
    }
}
