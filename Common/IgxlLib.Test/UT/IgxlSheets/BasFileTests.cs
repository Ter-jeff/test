using System.Collections.Generic;
using System.IO;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class BasFileTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void BasFile_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "BasFile";

            // Act
            var basFile = new BasFile(sheetName);

            // Assert
            Assert.IsNotNull(basFile);
            Assert.AreEqual(sheetName, basFile.Name);
            Assert.AreEqual("BasFile", basFile.SheetType);
            Assert.AreEqual("BasFile", basFile.IgxlSheetName);
            Assert.AreEqual(0, basFile.Rows.Count);
        }

        [TestMethod]
        public void BasFile_SheetType_IsCorrect()
        {
            // Arrange & Act
            var basFile = new BasFile("BasFile");

            // Assert
            Assert.AreEqual("BasFile", basFile.SheetType);
        }

        [TestMethod]
        public void BasFile_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var basFile = new BasFile("BasFile");

            // Assert
            Assert.AreEqual("BasFile", basFile.IgxlSheetName);
        }

        [TestMethod]
        public void BasFile_AddRow()
        {
            // Arrange
            var basFile = new BasFile("BasFile");
            var basRow = new BasRow { Text = "BasRow1" };

            // Act
            basFile.AddRow(basRow);

            // Assert
            Assert.AreEqual(1, basFile.Rows.Count);
            Assert.AreEqual("BasRow1", basFile.Rows[0].Text);
        }

        [TestMethod]
        public void BasFile_AddMultipleRows()
        {
            // Arrange
            var basFile = new BasFile("BasFile");
            var rows = new List<BasRow>
            {
                new() { Text = "BasRow1" },
                new() { Text = "BasRow2" },
                new() { Text = "BasRow3" }
            };

            // Act
            basFile.AddRows(rows);

            // Assert
            Assert.AreEqual(3, basFile.Rows.Count);
        }

        [TestMethod]
        public void BasFile_Write_CreatesFile()
        {
            // Arrange
            var basFile = new BasFile("BasFile");
            basFile.AddRow(new BasRow { Text = "Line 1" });
            basFile.AddRow(new BasRow { Text = "Line 2" });
            basFile.AddRow(new BasRow { Text = "Line 3" });

            string fileName = Path.Combine(Path.GetTempPath(), "test_basfile.txt");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            try
            {
                // Act
                basFile.Write(fileName);

                // Assert
                Assert.IsTrue(File.Exists(fileName));
                string[] lines = File.ReadAllLines(fileName);
                Assert.AreEqual(3, lines.Length);
                Assert.AreEqual("Line 1", lines[0]);
                Assert.AreEqual("Line 2", lines[1]);
                Assert.AreEqual("Line 3", lines[2]);
            }
            finally
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }
        }

        [TestMethod]
        public void BasFile_Write_WithVersion()
        {
            // Arrange
            var basFile = new BasFile("BasFile");
            basFile.AddRow(new BasRow { Text = "Test Line" });

            string fileName = Path.Combine(Path.GetTempPath(), "test_basfile_v2.txt");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            try
            {
                // Act
                basFile.Write(fileName, "2.0");

                // Assert
                Assert.IsTrue(File.Exists(fileName));
                string[] lines = File.ReadAllLines(fileName);
                Assert.AreEqual(1, lines.Length);
                Assert.AreEqual("Test Line", lines[0]);
            }
            finally
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }
        }

        [TestMethod]
        public void BasFile_Name_CanBeSet()
        {
            // Arrange
            var basFile = new BasFile("BasFile")
            {
                // Act
                Name = "NewBasFileName"
            };

            // Assert
            Assert.AreEqual("NewBasFileName", basFile.Name);
        }

        [TestMethod]
        public void BasFile_Rows_InitializedEmpty()
        {
            // Arrange & Act
            var basFile = new BasFile("BasFile");

            // Assert
            Assert.AreEqual(0, basFile.Rows.Count);
        }
    }
}
