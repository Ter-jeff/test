using System;
using System.Collections.Generic;
using System.IO;

using FileDiffLib;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class CharSheetTests
    {
        public string InputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        public string OutputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        public string ExpectPath = Path.Combine(Directory.GetCurrentDirectory(), "Expected");

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void CharSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "Characterization";

            // Act
            var charSheet = new CharSheet(sheetName);

            // Assert
            Assert.IsNotNull(charSheet);
            Assert.AreEqual(sheetName, charSheet.Name);
            Assert.AreEqual("DTCharacterizationSheet", charSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.Characterization, charSheet.IgxlSheetName);
            Assert.AreEqual(0, charSheet.Rows.Count);
        }

        [TestMethod]
        public void CharSheet_AddRow()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            var charSetup = new CharSetup
            {
                SetupName = "CharSetup1",
                TestMethod = "TestMethod1"
            };

            // Act
            charSheet.AddRow(charSetup);

            // Assert
            Assert.AreEqual(1, charSheet.Rows.Count);
            Assert.AreEqual("CharSetup1", charSheet.Rows[0].SetupName);
        }

        [TestMethod]
        public void CharSheet_AddRows()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            var rows = new List<CharSetup>
            {
                new() { SetupName = "Setup1" },
                new() { SetupName = "Setup2" },
                new() { SetupName = "Setup3" }
            };

            // Act
            charSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, charSheet.Rows.Count);
        }

        [TestMethod]
        public void CharSheet_RemoveRow()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            var row1 = new CharSetup { SetupName = "Setup1" };
            var row2 = new CharSetup { SetupName = "Setup2" };
            charSheet.AddRow(row1);
            charSheet.AddRow(row2);

            // Act
            charSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, charSheet.Rows.Count);
            Assert.AreEqual("Setup2", charSheet.Rows[0].SetupName);
        }

        [TestMethod]
        public void CharSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var charSheet = new CharSheet("Characterization");

            // Assert
            Assert.AreEqual("DTCharacterizationSheet", charSheet.SheetType);
        }

        [TestMethod]
        public void CharSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var charSheet = new CharSheet("Characterization");

            // Assert
            Assert.AreEqual(IgxlSheetNames.Characterization, charSheet.IgxlSheetName);
        }

        [TestMethod]
        public void CharSheet_ClearRows()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            charSheet.AddRow(new CharSetup { SetupName = "Setup1" });
            charSheet.AddRow(new CharSetup { SetupName = "Setup2" });
            Assert.AreEqual(2, charSheet.Rows.Count);

            // Act
            charSheet.Rows.Clear();

            // Assert
            Assert.AreEqual(0, charSheet.Rows.Count);
        }

        [TestMethod]
        public void CharSheet_Write_WithEmptyRows_DoesNotThrow()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"CharSheet_{Guid.NewGuid()}.txt");

            try
            {
                // Act
                charSheet.Write(tempFileName);

                // Assert - Verify the file was NEVER created
                Assert.IsFalse(File.Exists(tempFileName));
            }
            finally
            {
                // Clean up just in case the bug causes the file to be created
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }

        [TestMethod]
        public void CharSheet_Write_DoesNotThrow()
        {
            string subName = "CharSheet_Write";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            var charSheet = new CharSheet("Characterization");
            var rows = new List<CharSetup>();
            for (int i = 0; i < 5; i++)
            {
                var setup = new CharSetup { SetupName = $"Setup{i}" };
                rows.Add(setup);
                setup.AddStep(new CharStep("", "") { StepName = "Step1", VoltageType = "DC" });
                charSheet.AddRow(setup);
            }
            string file = Path.Combine(outputPath, $"CharSheet.txt");

            // Act
            charSheet.Write(file);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void CharSheet_Write_WithVersion_DoesNotThrow()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"CharSheet_{Guid.NewGuid()}.txt");

            try
            {
                // Act
                charSheet.Write(tempFileName, "2.0");

                // Assert - Verify the file was NEVER created
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
        public void CharSheet_AddMultipleRows_IncrementsCount()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");

            // Act
            for (int i = 0; i < 5; i++)
            {
                charSheet.AddRow(new CharSetup { SetupName = $"Setup{i}", TestMethod = $"Method{i}" });
            }

            // Assert
            Assert.AreEqual(5, charSheet.Rows.Count);
        }

        [TestMethod]
        public void CharSheet_RemoveMultipleRows()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            var rows = new List<CharSetup>();
            for (int i = 0; i < 5; i++)
            {
                var setup = new CharSetup { SetupName = $"Setup{i}" };
                rows.Add(setup);
                charSheet.AddRow(setup);
            }

            // Act
            charSheet.RemoveRow(rows[2]);
            charSheet.RemoveRow(rows[0]);

            // Assert
            Assert.AreEqual(3, charSheet.Rows.Count);
        }

        [TestMethod]
        public void CharSheet_AddRows_WithEmptyList()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            var emptyRows = new List<CharSetup>();

            // Act
            charSheet.AddRows(emptyRows);

            // Assert
            Assert.AreEqual(0, charSheet.Rows.Count);
        }

        [TestMethod]
        public void CharSheet_Name_CanBeSet()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization")
            {
                // Act
                Name = "CustomName"
            };

            // Assert
            Assert.AreEqual("CustomName", charSheet.Name);
        }

        [TestMethod]
        public void CharSheet_Rows_IsNotNull()
        {
            // Arrange & Act
            var charSheet = new CharSheet("Characterization");

            // Assert
            Assert.IsInstanceOfType(charSheet.Rows, typeof(List<CharSetup>));
        }

        [TestMethod]
        public void CharSheet_GetRow_RetrievesCorrectRow()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            var setup1 = new CharSetup { SetupName = "Setup1" };
            var setup2 = new CharSetup { SetupName = "Setup2" };
            charSheet.AddRow(setup1);
            charSheet.AddRow(setup2);

            // Act
            CharSetup retrieved = charSheet.Rows[0];

            // Assert
            Assert.AreEqual("Setup1", retrieved.SetupName);
        }

        [TestMethod]
        public void CharSheet_RemoveRow_WithNonExistentRow()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            var existingRow = new CharSetup { SetupName = "Setup1" };
            var nonExistentRow = new CharSetup { SetupName = "Setup2" };
            charSheet.AddRow(existingRow);

            // Act
            bool removed = charSheet.Rows.Remove(nonExistentRow);

            // Assert
            Assert.IsFalse(removed);
            Assert.AreEqual(1, charSheet.Rows.Count);
        }

        [TestMethod]
        public void CharSheet_Contains_ChecksForRow()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");
            var setup = new CharSetup { SetupName = "Setup1" };
            charSheet.AddRow(setup);

            // Act
            bool contains = charSheet.Rows.Contains(setup);

            // Assert
            Assert.IsTrue(contains);
        }

        [TestMethod]
        public void CharSheet_Multiple_Operations_InSequence()
        {
            // Arrange
            var charSheet = new CharSheet("Characterization");

            // Act
            charSheet.AddRow(new CharSetup { SetupName = "Setup1" });
            charSheet.AddRow(new CharSetup { SetupName = "Setup2" });
            charSheet.AddRow(new CharSetup { SetupName = "Setup3" });
            CharSetup removedRow = charSheet.Rows[1];
            charSheet.RemoveRow(removedRow);

            // Assert
            Assert.AreEqual(2, charSheet.Rows.Count);
            Assert.AreEqual("Setup1", charSheet.Rows[0].SetupName);
            Assert.AreEqual("Setup3", charSheet.Rows[1].SetupName);
        }
    }
}
