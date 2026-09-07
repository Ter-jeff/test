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
    public class MixedSignalSheetTests
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
        public void MixedSignalSheet_Write_DoesNotThrow()
        {
            string subName = "MixedSignalSheet_Write";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            var mixedSignalSheet = new MixedSignalSheet("LimitSets");
            var rows = new List<MixedSigRow>();
            for (int i = 0; i < 5; i++)
            {
                var row = new MixedSigRow
                {
                    Name = $"Setup{i}",
                };
                rows.Add(row);
                mixedSignalSheet.AddRow(row);
            }
            string file = Path.Combine(outputPath, $"MixedSignalSheet.txt");

            // Act
            mixedSignalSheet.Write(file);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void MixedSignalSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "MixedSignal";

            // Act
            var mixedSignalSheet = new MixedSignalSheet(sheetName);

            // Assert
            Assert.IsNotNull(mixedSignalSheet);
            Assert.AreEqual(sheetName, mixedSignalSheet.Name);
            Assert.AreEqual("DTMixedSignalTimingSheet", mixedSignalSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.MixedSignal, mixedSignalSheet.IgxlSheetName);
            Assert.AreEqual(0, mixedSignalSheet.Rows.Count);
        }

        [TestMethod]
        public void MixedSignalSheet_AddRow()
        {
            // Arrange
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal");
            var mixedSigRow = new MixedSigRow
            {
                ColumnA = "A",
                Name = "MixedSig1",
                Subset = "Subset1",
                Type = "Type1",
                Id = "ID1"
            };

            // Act
            mixedSignalSheet.AddRow(mixedSigRow);

            // Assert
            Assert.AreEqual(1, mixedSignalSheet.Rows.Count);
            Assert.AreEqual("MixedSig1", mixedSignalSheet.Rows[0].Name);
            Assert.AreEqual("Subset1", mixedSignalSheet.Rows[0].Subset);
        }

        [TestMethod]
        public void MixedSignalSheet_AddRows()
        {
            // Arrange
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal");
            var rows = new List<MixedSigRow>
            {
                new() { Name = "MixedSig1", Subset = "Subset1" },
                new() { Name = "MixedSig2", Subset = "Subset2" },
                new() { Name = "MixedSig3", Subset = "Subset3" }
            };

            // Act
            mixedSignalSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, mixedSignalSheet.Rows.Count);
        }

        [TestMethod]
        public void MixedSignalSheet_RemoveRow()
        {
            // Arrange
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal");
            var row1 = new MixedSigRow { Name = "MixedSig1", Subset = "Subset1" };
            var row2 = new MixedSigRow { Name = "MixedSig2", Subset = "Subset2" };
            mixedSignalSheet.AddRow(row1);
            mixedSignalSheet.AddRow(row2);

            // Act
            mixedSignalSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, mixedSignalSheet.Rows.Count);
            Assert.AreEqual("MixedSig2", mixedSignalSheet.Rows[0].Name);
        }

        [TestMethod]
        public void MixedSignalSheet_InsertRow()
        {
            // Arrange
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal");
            var row1 = new MixedSigRow { Name = "MixedSig1", Subset = "Subset1" };
            var row3 = new MixedSigRow { Name = "MixedSig3", Subset = "Subset3" };
            var rowToInsert = new MixedSigRow { Name = "MixedSig2", Subset = "Subset2" };
            mixedSignalSheet.AddRow(row1);
            mixedSignalSheet.AddRow(row3);

            // Act
            int index = mixedSignalSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, mixedSignalSheet.Rows.Count);
            Assert.AreEqual("MixedSig2", mixedSignalSheet.Rows[1].Name);
            Assert.AreEqual("MixedSig3", mixedSignalSheet.Rows[2].Name);
        }

        [TestMethod]
        public void MixedSignalSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal");

            // Assert
            Assert.AreEqual("DTMixedSignalTimingSheet", mixedSignalSheet.SheetType);
        }

        [TestMethod]
        public void MixedSignalSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal");

            // Assert
            Assert.AreEqual(IgxlSheetNames.MixedSignal, mixedSignalSheet.IgxlSheetName);
        }

        [TestMethod]
        public void MixedSignalSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal");

            // Assert
            Assert.IsNotNull(mixedSignalSheet.GetErrors());
            Assert.AreEqual(0, mixedSignalSheet.GetErrors().Count);
        }

        [TestMethod]
        public void MixedSignalSheet_Name_CanBeSet()
        {
            // Arrange
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal")
            {
                // Act
                Name = "NewMixedSignalName"
            };

            // Assert
            Assert.AreEqual("NewMixedSignalName", mixedSignalSheet.Name);
        }

        [TestMethod]
        public void MixedSignalSheet_MixedSigRow_WithAllProperties()
        {
            // Arrange
            var mixedSignalSheet = new MixedSignalSheet("MixedSignal");
            var mixedSigRow = new MixedSigRow
            {
                ColumnA = "A",
                Name = "MixedSig1",
                Subset = "Subset1",
                Type = "Type1",
                Id = "ID1",
                Fs = "100",
                N = "10",
                Fr = "50",
                M = "5",
                Usr = "User1",
                Data = "Data1",
                Definition = "Def1",
                Filter = "Filter1",
                Settings = "Settings1",
                WaveName = "Wave1",
                Amplitude = "1.0",
                Offset = "0.5",
                OldInstData = "OldData",
                Comment = "Test comment"
            };

            // Act
            mixedSignalSheet.AddRow(mixedSigRow);

            // Assert
            Assert.AreEqual(1, mixedSignalSheet.Rows.Count);
            Assert.AreEqual("MixedSig1", mixedSignalSheet.Rows[0].Name);
            Assert.AreEqual("Subset1", mixedSignalSheet.Rows[0].Subset);
            Assert.AreEqual("100", mixedSignalSheet.Rows[0].Fs);
            Assert.AreEqual("Wave1", mixedSignalSheet.Rows[0].WaveName);
        }
    }
}
