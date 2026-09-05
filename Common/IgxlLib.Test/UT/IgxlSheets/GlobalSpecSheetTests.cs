using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class GlobalSpecSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void GlobalSpecSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "GlobalSpec";

            // Act
            var globalSpecSheet = new GlobalSpecSheet(sheetName);

            // Assert
            Assert.IsNotNull(globalSpecSheet);
            Assert.AreEqual(sheetName, globalSpecSheet.Name);
            Assert.AreEqual("DTGlobalSpecSheet", globalSpecSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.GlobalSpec, globalSpecSheet.IgxlSheetName);
            // Should have default rows (Vcl_default, Vch_default, Vph_default)
            Assert.AreEqual(3, globalSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void GlobalSpecSheet_Constructor_WithSheetName_NoDefaults()
        {
            // Arrange
            string sheetName = "GlobalSpec";

            // Act
            var globalSpecSheet = new GlobalSpecSheet(sheetName, isAddDefault: false);

            // Assert
            Assert.IsNotNull(globalSpecSheet);
            Assert.AreEqual(sheetName, globalSpecSheet.Name);
            Assert.AreEqual("DTGlobalSpecSheet", globalSpecSheet.SheetType);
            Assert.AreEqual(0, globalSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void GlobalSpecSheet_AddRow()
        {
            // Arrange
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec", isAddDefault: false);
            var globalSpec = new GlobalSpec("TestSpec", "Job1", "1.5", "Test comment");

            // Act
            globalSpecSheet.AddRow(globalSpec);

            // Assert
            Assert.AreEqual(1, globalSpecSheet.Rows.Count);
            Assert.AreEqual("TestSpec", globalSpecSheet.Rows[0].Symbol);
            Assert.AreEqual("Job1", globalSpecSheet.Rows[0].Job);
            Assert.AreEqual("1.5", globalSpecSheet.Rows[0].Value);
        }

        [TestMethod]
        public void GlobalSpecSheet_AddRow_DuplicateSymbolAndJob_NotAdded()
        {
            // Arrange
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec", isAddDefault: false);
            var globalSpec1 = new GlobalSpec("TestSpec", "Job1", "1.5", "Test comment");
            var globalSpec2 = new GlobalSpec("TestSpec", "Job2", "2.0", "Another test");
            var globalSpec3 = new GlobalSpec("TestSpec", "Job2", "2.5", "Another test2");

            // Act
            globalSpecSheet.AddRow(globalSpec1);
            globalSpecSheet.AddRow(globalSpec2);
            globalSpecSheet.AddRow(globalSpec3);

            // Assert
            Assert.AreEqual(2, globalSpecSheet.Rows.Count);
            Assert.AreEqual("Job1", globalSpecSheet.Rows[0].Job);
            Assert.AreEqual("2.0", globalSpecSheet.Rows[1].Value);
        }

        [TestMethod]
        public void GlobalSpecSheet_AddRow_CaseInsensitiveSymbol_SameJob()
        {
            // Arrange
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec", isAddDefault: false);
            var globalSpec1 = new GlobalSpec("testspec", "Job1", "1.5", "Test comment");
            var globalSpec2 = new GlobalSpec("TESTSPEC", "Job1", "2.0", "Another test");

            // Act
            globalSpecSheet.AddRow(globalSpec1);
            globalSpecSheet.AddRow(globalSpec2);

            // Assert
            Assert.AreEqual(1, globalSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void GlobalSpecSheet_AddRow_CaseInsensitiveSymbol_DifferentJob()
        {
            // Arrange
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec", isAddDefault: false);
            var globalSpec1 = new GlobalSpec("testspec", "Job1", "1.5", "Test comment");
            var globalSpec2 = new GlobalSpec("TESTSPEC", "Job2", "2.0", "Another test");

            // Act
            globalSpecSheet.AddRow(globalSpec1);
            globalSpecSheet.AddRow(globalSpec2);

            // Assert
            Assert.AreEqual(2, globalSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void GlobalSpecSheet_AddRows()
        {
            // Arrange
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec", isAddDefault: false);
            var rows = new List<GlobalSpec>
            {
                new("Spec1", "Job1", "1.0", "Comment 1"),
                new("Spec2", "Job2", "2.0", "Comment 2"),
                new("Spec3", "Job3", "3.0", "Comment 3")
            };

            // Act
            globalSpecSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, globalSpecSheet.Rows.Count);
        }

        [TestMethod]
        public void GlobalSpecSheet_RemoveRow()
        {
            // Arrange
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec", isAddDefault: false);
            var row1 = new GlobalSpec("Spec1", "Job1", "1.0", "Comment 1");
            var row2 = new GlobalSpec("Spec2", "Job2", "2.0", "Comment 2");
            globalSpecSheet.AddRow(row1);
            globalSpecSheet.AddRow(row2);

            // Act
            globalSpecSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, globalSpecSheet.Rows.Count);
            Assert.AreEqual("Spec2", globalSpecSheet.Rows[0].Symbol);
        }

        [TestMethod]
        public void GlobalSpecSheet_FindRowList_BothSymbolAndJob()
        {
            // Arrange
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec", isAddDefault: false);
            globalSpecSheet.AddRow(new GlobalSpec("Spec1", "JobA", "1.0", "Comment 1"));
            globalSpecSheet.AddRow(new GlobalSpec("Spec1", "JobB", "1.5", "Comment 2"));
            globalSpecSheet.AddRow(new GlobalSpec("Spec2", "JobA", "2.0", "Comment 3"));

            // Act
            List<GlobalSpec> results = globalSpecSheet.FindRowList("Spec1", "JobA");

            // Assert
            Assert.AreEqual(1, results.Count);
            Assert.AreEqual("Spec1", results[0].Symbol);
            Assert.AreEqual("JobA", results[0].Job);
        }

        [TestMethod]
        public void GlobalSpecSheet_FindRowList_NoMatches()
        {
            // Arrange
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec", isAddDefault: false);
            globalSpecSheet.AddRow(new GlobalSpec("Spec1", "JobA", "1.0", "Comment 1"));

            // Act
            List<GlobalSpec> results = globalSpecSheet.FindRowList("NonExistent", "JobX");

            // Assert
            Assert.AreEqual(0, results.Count);
        }

        [TestMethod]
        public void GlobalSpecSheet_DefaultRows_VclDefault()
        {
            // Arrange & Act
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec");

            // Assert
            List<GlobalSpec> vclDefault = globalSpecSheet.FindRowList("Vcl_default", "");
            Assert.AreEqual(1, vclDefault.Count);
            Assert.AreEqual("=-1", vclDefault[0].Value);
        }

        [TestMethod]
        public void GlobalSpecSheet_DefaultRows_VchDefault()
        {
            // Arrange & Act
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec");

            // Assert
            List<GlobalSpec> vchDefault = globalSpecSheet.FindRowList("Vch_default", "");
            Assert.AreEqual(1, vchDefault.Count);
            Assert.AreEqual("=6", vchDefault[0].Value);
        }

        [TestMethod]
        public void GlobalSpecSheet_DefaultRows_VphDefault()
        {
            // Arrange & Act
            var globalSpecSheet = new GlobalSpecSheet("GlobalSpec");

            // Assert
            List<GlobalSpec> vphDefault = globalSpecSheet.FindRowList("Vph_default", "");
            Assert.AreEqual(1, vphDefault.Count);
            Assert.AreEqual("=5", vphDefault[0].Value);
        }
    }
}
