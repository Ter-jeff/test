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
    public class LimitSetsSheetTests
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
        public void LimitSetsSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "LimitSets";

            // Act
            var limitSetsSheet = new LimitSetsSheet(sheetName);

            // Assert
            Assert.IsNotNull(limitSetsSheet);
            Assert.AreEqual(sheetName, limitSetsSheet.Name);
            Assert.AreEqual("LimitSetsSheet", limitSetsSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.LimitSetsSheet, limitSetsSheet.IgxlSheetName);
        }

        [TestMethod]
        public void LimitSetsSheet_AddRow()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets");
            var limitRow = new LimitSetsRow();

            // Act
            limitSetsSheet.AddRow(limitRow);

            // Assert
            Assert.AreEqual(1, limitSetsSheet.Rows.Count);
        }

        [TestMethod]
        public void LimitSetsSheet_AddRows()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets");
            var rows = new List<LimitSetsRow>
            {
                new(),
                new(),
                new()
            };

            // Act
            limitSetsSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, limitSetsSheet.Rows.Count);
        }

        [TestMethod]
        public void LimitSetsSheet_RemoveRow()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets");
            var row1 = new LimitSetsRow();
            var row2 = new LimitSetsRow() { TestName = "Limit2" };
            limitSetsSheet.AddRow(row1);
            limitSetsSheet.AddRow(row2);

            // Act
            limitSetsSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, limitSetsSheet.Rows.Count);
            Assert.AreEqual("Limit2", limitSetsSheet.Rows[0].TestName);
        }

        [TestMethod]
        public void LimitSetsSheet_InsertRow()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets");
            var row1 = new LimitSetsRow();
            var row2 = new LimitSetsRow() { TestName = "Limit3" };
            var rowToInsert = new LimitSetsRow();
            limitSetsSheet.AddRow(row1);
            limitSetsSheet.AddRow(row2);

            // Act
            int index = limitSetsSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, limitSetsSheet.Rows.Count);
            Assert.AreEqual("Limit3", limitSetsSheet.Rows[2].TestName);
        }

        [TestMethod]
        public void LimitSetsSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var limitSetsSheet = new LimitSetsSheet("LimitSets");

            // Assert
            Assert.AreEqual("LimitSetsSheet", limitSetsSheet.SheetType);
        }

        [TestMethod]
        public void LimitSetsSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var limitSetsSheet = new LimitSetsSheet("LimitSets");

            // Assert
            Assert.IsNotNull(limitSetsSheet.GetErrors());
            Assert.AreEqual(0, limitSetsSheet.GetErrors().Count);
        }

        [TestMethod]
        public void LimitSetsSheet_Name_CanBeSet()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets")
            {
                // Act
                Name = "NewLimitSetsName"
            };

            // Assert
            Assert.AreEqual("NewLimitSetsName", limitSetsSheet.Name);
        }

        [TestMethod]
        public void LimitSetsSheet_Write_WithEmptyRows_DoesNotCreateFile()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"LimitSetsSheet_Empty_{Guid.NewGuid()}.igx");

            try
            {
                // Act
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
                limitSetsSheet.Write(tempFileName);

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
        public void LimitSetsSheet_Write_DoesNotThrow()
        {
            string subName = "LimitSetsSheet_Write";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets");
            var rows = new List<LimitSetsRow>();
            for (int i = 0; i < 5; i++)
            {
                var row = new LimitSetsRow
                {
                    TestName = $"Setup{i}",
                    LimitSetsCat = [new() { CategoryName = "CategoryName" }]
                };
                rows.Add(row);
                limitSetsSheet.AddRow(row);
            }
            string file = Path.Combine(outputPath, $"LimitSetsSheet.txt");

            // Act
            limitSetsSheet.Write(file);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void LimitSetsSheet_Write_WithVersion_DoesNotThrow()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"LimitSetsSheet_Version_{Guid.NewGuid()}.igx");

            try
            {
                // Act & Assert - should not throw
                limitSetsSheet.Write(tempFileName, "2.0");
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
        public void LimitSetsSheet_WriteContent_WithEmptyRows_ReturnsNull()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets");

            // Act
            System.Text.StringBuilder content = limitSetsSheet.WriteContent("2.0");

            // Assert
            Assert.IsNull(content);
        }

        [TestMethod]
        public void LimitSetsSheet_Constructor_WithValidSheetName()
        {
            // Arrange & Act
            var limitSetsSheet = new LimitSetsSheet("LimitSets");

            // Assert
            Assert.IsNotNull(limitSetsSheet);
            Assert.AreEqual("LimitSets", limitSetsSheet.Name);
        }

        [TestMethod]
        public void LimitSetsSheet_Constructor_WithLongSheetName_ThrowsException()
        {
            // Arrange
            string longSheetName = new('A', 32);

            // Act
            Assert.ThrowsException<Exception>(() => new LimitSetsSheet(longSheetName));
        }

        [TestMethod]
        public void LimitSetsSheet_SourceSheet_CanBeSet()
        {
            // Arrange
            var limitSetsSheet = new LimitSetsSheet("LimitSets")
            {
                // Act
                SourceSheet = "SourceSheet1"
            };

            // Assert
            Assert.AreEqual("SourceSheet1", limitSetsSheet.SourceSheet);
        }
    }
}
