using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class ReferenceSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void ReferenceSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "Reference";

            // Act
            var referenceSheet = new ReferenceSheet(sheetName);

            // Assert
            Assert.IsNotNull(referenceSheet);
            Assert.AreEqual(sheetName, referenceSheet.Name);
            Assert.AreEqual("DTReferencesSheet", referenceSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.Reference, referenceSheet.IgxlSheetName);
        }

        [TestMethod]
        public void ReferenceSheet_AddRow()
        {
            // Arrange
            var referenceSheet = new ReferenceSheet("Reference");
            var referenceRow = new ReferenceRow
            {
                FilePath = "Reference1",
                Comment = "RefType1"
            };

            // Act
            referenceSheet.AddRow(referenceRow);

            // Assert
            Assert.AreEqual(1, referenceSheet.Rows.Count);
            Assert.AreEqual("Reference1", referenceSheet.Rows[0].FilePath);
            Assert.AreEqual("RefType1", referenceSheet.Rows[0].Comment);
        }

        [TestMethod]
        public void ReferenceSheet_AddRows()
        {
            // Arrange
            var referenceSheet = new ReferenceSheet("Reference");
            var rows = new List<ReferenceRow>
            {
                new() { FilePath = "Reference1", Comment ="Type1" },
                new() { FilePath = "Reference2", Comment ="Type2" },
                new() { FilePath = "Reference3", Comment ="Type3" }
            };

            // Act
            referenceSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, referenceSheet.Rows.Count);
        }

        [TestMethod]
        public void ReferenceSheet_RemoveRow()
        {
            // Arrange
            var referenceSheet = new ReferenceSheet("Reference");
            var row1 = new ReferenceRow { FilePath = "Reference1", Comment = "Type1" };
            var row2 = new ReferenceRow { FilePath = "Reference2", Comment = "Type2" };
            referenceSheet.AddRow(row1);
            referenceSheet.AddRow(row2);

            // Act
            referenceSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, referenceSheet.Rows.Count);
            Assert.AreEqual("Reference2", referenceSheet.Rows[0].FilePath);
        }

        [TestMethod]
        public void ReferenceSheet_InsertRow()
        {
            // Arrange
            var referenceSheet = new ReferenceSheet("Reference");
            var row1 = new ReferenceRow { FilePath = "Reference1", Comment = "Type1" };
            var row3 = new ReferenceRow { FilePath = "Reference3", Comment = "Type3" };
            var rowToInsert = new ReferenceRow { FilePath = "Reference2", Comment = "Type2" };
            referenceSheet.AddRow(row1);
            referenceSheet.AddRow(row3);

            // Act
            int index = referenceSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, referenceSheet.Rows.Count);
            Assert.AreEqual("Reference2", referenceSheet.Rows[1].FilePath);
            Assert.AreEqual("Reference3", referenceSheet.Rows[2].FilePath);
        }

        [TestMethod]
        public void ReferenceSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var referenceSheet = new ReferenceSheet("Reference");

            // Assert
            Assert.AreEqual("DTReferencesSheet", referenceSheet.SheetType);
        }

        [TestMethod]
        public void ReferenceSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var referenceSheet = new ReferenceSheet("Reference");

            // Assert
            Assert.AreEqual(IgxlSheetNames.Reference, referenceSheet.IgxlSheetName);
        }

        [TestMethod]
        public void ReferenceSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var referenceSheet = new ReferenceSheet("Reference");

            // Assert
            Assert.IsNotNull(referenceSheet.GetErrors());
            Assert.AreEqual(0, referenceSheet.GetErrors().Count);
        }

        [TestMethod]
        public void ReferenceSheet_Name_CanBeSet()
        {
            // Arrange
            var referenceSheet = new ReferenceSheet("Reference")
            {
                // Act
                Name = "NewReferenceName"
            };

            // Assert
            Assert.AreEqual("NewReferenceName", referenceSheet.Name);
        }

        [TestMethod]
        public void ReferenceSheet_ReferenceRow_WithProperties()
        {
            // Arrange
            var referenceSheet = new ReferenceSheet("Reference");
            var referenceRow = new ReferenceRow
            {
                FilePath = "Reference1",
                Comment = "RefType1",
            };

            // Act
            referenceSheet.AddRow(referenceRow);

            // Assert
            Assert.AreEqual(1, referenceSheet.Rows.Count);
            Assert.AreEqual("Reference1", referenceSheet.Rows[0].FilePath);
            Assert.AreEqual("RefType1", referenceSheet.Rows[0].Comment);
        }
    }
}
