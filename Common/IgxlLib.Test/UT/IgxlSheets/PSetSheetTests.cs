using System;
using System.Collections.Generic;
using System.IO;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class PSetSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void PSetSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "PSet";

            // Act
            var psetSheet = new PSetSheet(sheetName);

            // Assert
            Assert.IsNotNull(psetSheet);
            Assert.AreEqual(sheetName, psetSheet.Name);
            Assert.AreEqual("DTPsetsSheet", psetSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.PSet, psetSheet.IgxlSheetName);
            Assert.AreEqual(0, psetSheet.Rows.Count);
        }

        [TestMethod]
        public void PSetSheet_AddRow()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            var pset = new PSet
            {
                Name = "PSet1",
                Pin = "Pin1",
                InstrumentType = "DCVoltage"
            };

            // Act
            psetSheet.AddRow(pset);

            // Assert
            Assert.AreEqual(1, psetSheet.Rows.Count);
            Assert.AreEqual("PSet1", psetSheet.Rows[0].Name);
            Assert.AreEqual("Pin1", psetSheet.Rows[0].Pin);
        }

        [TestMethod]
        public void PSetSheet_AddRows()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            var rows = new List<PSet>
            {
                new() { Name = "PSet1", Pin = "Pin1" },
                new() { Name = "PSet2", Pin = "Pin2" },
                new() { Name = "PSet3", Pin = "Pin3" }
            };

            // Act
            psetSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, psetSheet.Rows.Count);
        }

        [TestMethod]
        public void PSetSheet_RemoveRow()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            var row1 = new PSet { Name = "PSet1", Pin = "Pin1" };
            var row2 = new PSet { Name = "PSet2", Pin = "Pin2" };
            psetSheet.AddRow(row1);
            psetSheet.AddRow(row2);

            // Act
            psetSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, psetSheet.Rows.Count);
            Assert.AreEqual("PSet2", psetSheet.Rows[0].Name);
        }

        [TestMethod]
        public void PSetSheet_InsertRow()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            var row1 = new PSet { Name = "PSet1", Pin = "Pin1" };
            var row3 = new PSet { Name = "PSet3", Pin = "Pin3" };
            var rowToInsert = new PSet { Name = "PSet2", Pin = "Pin2" };
            psetSheet.AddRow(row1);
            psetSheet.AddRow(row3);

            // Act
            int index = psetSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, psetSheet.Rows.Count);
            Assert.AreEqual("PSet2", psetSheet.Rows[1].Name);
            Assert.AreEqual("PSet3", psetSheet.Rows[2].Name);
        }

        [TestMethod]
        public void PSetSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var psetSheet = new PSetSheet("PSet");

            // Assert
            Assert.AreEqual("DTPsetsSheet", psetSheet.SheetType);
        }

        [TestMethod]
        public void PSetSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var psetSheet = new PSetSheet("PSet");

            // Assert
            Assert.AreEqual(IgxlSheetNames.PSet, psetSheet.IgxlSheetName);
        }

        [TestMethod]
        public void PSetSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var psetSheet = new PSetSheet("PSet");

            // Assert
            Assert.IsNotNull(psetSheet.GetErrors());
            Assert.AreEqual(0, psetSheet.GetErrors().Count);
        }

        [TestMethod]
        public void PSetSheet_Name_CanBeSet()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet")
            {
                // Act
                Name = "NewPSetName"
            };

            // Assert
            Assert.AreEqual("NewPSetName", psetSheet.Name);
        }

        [TestMethod]
        public void PSetSheet_PSet_WithAllProperties()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            var pset = new PSet
            {
                Name = "PSet1",
                Pin = "Pin1",
                InstrumentType = "DCVoltage",
                Comment = "Test comment"
            };

            // Act
            psetSheet.AddRow(pset);

            // Assert
            Assert.AreEqual(1, psetSheet.Rows.Count);
            Assert.AreEqual("PSet1", psetSheet.Rows[0].Name);
            Assert.AreEqual("Pin1", psetSheet.Rows[0].Pin);
            Assert.AreEqual("DCVoltage", psetSheet.Rows[0].InstrumentType);
            Assert.AreEqual("Test comment", psetSheet.Rows[0].Comment);
        }

        [TestMethod]
        public void PSetSheet_Write_WithEmptyRows_DoesNotCreateFile()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"PSetSheet_Empty_{Guid.NewGuid()}.igx");

            try
            {
                // Act
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
                psetSheet.Write(tempFileName);

                // Assert
                Assert.IsFalse(File.Exists(tempFileName));
            }
            finally
            {
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }

        [TestMethod]
        public void PSetSheet_Write_WithRows_DoesNotThrow()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            var pset = new PSet
            {
                Name = "PSet1",
                Pin = "Pin1",
                InstrumentType = "DCVoltage",
                Comment = "Test PSet"
            };
            psetSheet.AddRow(pset);
            string tempFileName = Path.Combine(Path.GetTempPath(), $"PSetSheet_Data_{Guid.NewGuid()}.igx");

            try
            {
                // Act & Assert - should not throw
                psetSheet.Write(tempFileName);
            }
            finally
            {
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }

        [TestMethod]
        public void PSetSheet_Write_WithVersion_DoesNotThrow()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            var pset = new PSet { Name = "PSet1", Pin = "Pin1" };
            psetSheet.AddRow(pset);
            string tempFileName = Path.Combine(Path.GetTempPath(), $"PSetSheet_Version_{Guid.NewGuid()}.igx");

            try
            {
                // Act & Assert - should not throw
                psetSheet.Write(tempFileName, "2.0");
            }
            finally
            {
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }

        [TestMethod]
        public void PSetSheet_Write_MultipleRows_DoesNotThrow()
        {
            // Arrange
            var psetSheet = new PSetSheet("PSet");
            var rows = new List<PSet>
            {
                new() { Name = "PSet1", Pin = "Pin1" },
                new() { Name = "PSet2", Pin = "Pin2" },
                new() { Name = "PSet3", Pin = "Pin3" }
            };
            psetSheet.AddRows(rows);
            string tempFileName = Path.Combine(Path.GetTempPath(), $"PSetSheet_Multiple_{Guid.NewGuid()}.igx");

            try
            {
                // Act & Assert - should not throw
                psetSheet.Write(tempFileName);
            }
            finally
            {
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }
    }
}
