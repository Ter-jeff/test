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
    public class PatSetSubSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void PatSetSubSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "PatternSubroutine";

            // Act
            var patSetSubSheet = new PatSetSubSheet(sheetName);

            // Assert
            Assert.IsNotNull(patSetSubSheet);
            Assert.AreEqual(sheetName, patSetSubSheet.Name);
            Assert.AreEqual("DTPatternSubroutineSheet", patSetSubSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.PatternSubroutine, patSetSubSheet.IgxlSheetName);
        }

        [TestMethod]
        public void PatSetSubSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");

            // Assert
            Assert.AreEqual("DTPatternSubroutineSheet", patSetSubSheet.SheetType);
        }

        [TestMethod]
        public void PatSetSubSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");

            // Assert
            Assert.AreEqual(IgxlSheetNames.PatternSubroutine, patSetSubSheet.IgxlSheetName);
        }

        [TestMethod]
        public void PatSetSubSheet_AddRow()
        {
            // Arrange
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");
            var patSetSubRow = new PatSetSubRow { PatternFileName = "pattern1.txt" };

            // Act
            patSetSubSheet.AddRow(patSetSubRow);

            // Assert
            Assert.AreEqual(1, patSetSubSheet.Rows.Count);
            Assert.AreEqual("pattern1.txt", patSetSubSheet.Rows[0].PatternFileName);
        }

        [TestMethod]
        public void PatSetSubSheet_AddMultipleRows()
        {
            // Arrange
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");
            var rows = new List<PatSetSubRow>
            {
                new() { PatternFileName = "pattern1.txt" },
                new() { PatternFileName = "pattern2.txt" },
                new() { PatternFileName = "pattern3.txt" }
            };

            // Act
            patSetSubSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, patSetSubSheet.Rows.Count);
        }

        [TestMethod]
        public void PatSetSubSheet_GetExist_FindsExistingPattern()
        {
            // Arrange
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");
            var row1 = new PatSetSubRow { PatternFileName = "pattern1.txt", IsBackup = false };
            var row2 = new PatSetSubRow { PatternFileName = "pattern2.txt", IsBackup = false };
            patSetSubSheet.AddRow(row1);
            patSetSubSheet.AddRow(row2);

            // Act
            PatSetSubRow foundRow = patSetSubSheet.GetExist("pattern1.txt");

            // Assert
            Assert.IsNotNull(foundRow);
            Assert.AreEqual("pattern1.txt", foundRow.PatternFileName);
        }

        [TestMethod]
        public void PatSetSubSheet_GetExist_IgnoresBackupPatterns()
        {
            // Arrange
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");
            var backupRow = new PatSetSubRow { PatternFileName = "pattern.txt", IsBackup = true };
            var normalRow = new PatSetSubRow { PatternFileName = "pattern.txt", IsBackup = false };
            patSetSubSheet.AddRow(backupRow);
            patSetSubSheet.AddRow(normalRow);

            // Act
            PatSetSubRow foundRow = patSetSubSheet.GetExist("pattern.txt");

            // Assert
            Assert.IsNotNull(foundRow);
            Assert.IsFalse(foundRow.IsBackup);
        }

        [TestMethod]
        public void PatSetSubSheet_GetExist_ReturnsNullIfNotFound()
        {
            // Arrange
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");
            patSetSubSheet.AddRow(new PatSetSubRow { PatternFileName = "pattern1.txt" });

            // Act
            PatSetSubRow foundRow = patSetSubSheet.GetExist("nonexistent.txt");

            // Assert
            Assert.IsNull(foundRow);
        }

        [TestMethod]
        public void PatSetSubSheet_GetExist_IsCaseInsensitive()
        {
            // Arrange
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");
            var row = new PatSetSubRow { PatternFileName = "Pattern1.txt", IsBackup = false };
            patSetSubSheet.AddRow(row);

            // Act
            PatSetSubRow foundRow = patSetSubSheet.GetExist("pattern1.txt");

            // Assert
            Assert.IsNotNull(foundRow);
        }

        [TestMethod]
        public void PatSetSubSheet_Write_CreatesFile()
        {
            // Arrange
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");
            patSetSubSheet.AddRow(new PatSetSubRow { PatternFileName = "pattern1.txt" });

            string fileName = Path.Combine(Path.GetTempPath(), "test_patsetsubsheet.txt");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            try
            {
                // Act
                patSetSubSheet.Write(fileName);

                // Assert
                Assert.IsTrue(File.Exists(fileName));
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
        public void PatSetSubSheet_Name_CanBeSet()
        {
            // Arrange
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine")
            {
                // Act
                Name = "NewPatternSubroutine"
            };

            // Assert
            Assert.AreEqual("NewPatternSubroutine", patSetSubSheet.Name);
        }

        [TestMethod]
        public void PatSetSubSheet_Rows_InitializedEmpty()
        {
            // Arrange & Act
            var patSetSubSheet = new PatSetSubSheet("PatternSubroutine");

            // Assert
            Assert.AreEqual(0, patSetSubSheet.Rows.Count);
        }
    }
}
