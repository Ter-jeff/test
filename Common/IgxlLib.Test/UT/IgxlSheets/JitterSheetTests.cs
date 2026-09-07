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
    public class JitterSheetTests
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
        public void JitterSheet_Write_DoesNotThrow()
        {
            string subName = "JitterSheet_Write";
            string outputPath = Path.Combine(OutputPath, subName);
            string expectPath = Path.Combine(ExpectPath, subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            var jitterSheet = new JitterSheet("LimitSets");
            var rows = new List<JitterRow>();
            for (int i = 0; i < 5; i++)
            {
                var row = new JitterRow
                {
                    JitterSet = $"Setup{i}",
                };
                rows.Add(row);
                jitterSheet.AddRow(row);
            }
            string file = Path.Combine(outputPath, $"JitterSheet.txt");

            // Act
            jitterSheet.Write(file);

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }

        [TestMethod]
        public void JitterSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "Jitter";

            // Act
            var jitterSheet = new JitterSheet(sheetName);

            // Assert
            Assert.IsNotNull(jitterSheet);
            Assert.AreEqual(sheetName, jitterSheet.Name);
            Assert.AreEqual("DTSerialJitterSheet", jitterSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.JitterSheet, jitterSheet.IgxlSheetName);
            Assert.AreEqual(0, jitterSheet.Rows.Count);
        }

        [TestMethod]
        public void JitterSheet_AddRow()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            var jitterRow = new JitterRow
            {
                JitterSet = "JitterSet1",
                PinOrGroup = "Pin1",
                Mode = "Mode1",
                Period = "Period1",
                AmplitudeInUi = "1.0"
            };

            // Act
            jitterSheet.AddRow(jitterRow);

            // Assert
            Assert.AreEqual(1, jitterSheet.Rows.Count);
            Assert.AreEqual("JitterSet1", jitterSheet.Rows[0].JitterSet);
            Assert.AreEqual("Pin1", jitterSheet.Rows[0].PinOrGroup);
        }

        [TestMethod]
        public void JitterSheet_AddRows()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            var rows = new List<JitterRow>
            {
                new() { JitterSet = "JitterSet1", PinOrGroup = "Pin1" },
                new() { JitterSet = "JitterSet2", PinOrGroup = "Pin2" },
                new() { JitterSet = "JitterSet3", PinOrGroup = "Pin3" }
            };

            // Act
            jitterSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, jitterSheet.Rows.Count);
        }

        [TestMethod]
        public void JitterSheet_RemoveRow()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            var row1 = new JitterRow { JitterSet = "JitterSet1", PinOrGroup = "Pin1" };
            var row2 = new JitterRow { JitterSet = "JitterSet2", PinOrGroup = "Pin2" };
            jitterSheet.AddRow(row1);
            jitterSheet.AddRow(row2);

            // Act
            jitterSheet.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, jitterSheet.Rows.Count);
            Assert.AreEqual("JitterSet2", jitterSheet.Rows[0].JitterSet);
        }

        [TestMethod]
        public void JitterSheet_InsertRow()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            var row1 = new JitterRow { JitterSet = "JitterSet1", PinOrGroup = "Pin1" };
            var row3 = new JitterRow { JitterSet = "JitterSet3", PinOrGroup = "Pin3" };
            var rowToInsert = new JitterRow { JitterSet = "JitterSet2", PinOrGroup = "Pin2" };
            jitterSheet.AddRow(row1);
            jitterSheet.AddRow(row3);

            // Act
            int index = jitterSheet.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, jitterSheet.Rows.Count);
            Assert.AreEqual("JitterSet2", jitterSheet.Rows[1].JitterSet);
            Assert.AreEqual("JitterSet3", jitterSheet.Rows[2].JitterSet);
        }

        [TestMethod]
        public void JitterSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var jitterSheet = new JitterSheet("Jitter");

            // Assert
            Assert.AreEqual("DTSerialJitterSheet", jitterSheet.SheetType);
        }

        [TestMethod]
        public void JitterSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var jitterSheet = new JitterSheet("Jitter");

            // Assert
            Assert.AreEqual(IgxlSheetNames.JitterSheet, jitterSheet.IgxlSheetName);
        }

        [TestMethod]
        public void JitterSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var jitterSheet = new JitterSheet("Jitter");

            // Assert
            Assert.IsNotNull(jitterSheet.GetErrors());
            Assert.AreEqual(0, jitterSheet.GetErrors().Count);
        }

        [TestMethod]
        public void JitterSheet_Name_CanBeSet()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter")
            {
                // Act
                Name = "NewJitterName"
            };

            // Assert
            Assert.AreEqual("NewJitterName", jitterSheet.Name);
        }

        [TestMethod]
        public void JitterSheet_JitterRow_AllProperties()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            var jitterRow = new JitterRow
            {
                JitterSet = "JitterSet1",
                PinOrGroup = "Pin1",
                Mode = "Mode1",
                Period = "Period1",
                AmplitudeInUi = "1.0",
                SampleType = "SampleType1",
                TargetResolution = "0.1",
                TargetPatternReps = "100",
                BitsPerPattern = "10",
                MinPatternSkips = "5",
                BitPeriod = "1.5",
                Comment = "Test comment"
            };

            // Act
            jitterSheet.AddRow(jitterRow);

            // Assert
            Assert.AreEqual(1, jitterSheet.Rows.Count);
            Assert.AreEqual("JitterSet1", jitterSheet.Rows[0].JitterSet);
            Assert.AreEqual("Pin1", jitterSheet.Rows[0].PinOrGroup);
            Assert.AreEqual("Mode1", jitterSheet.Rows[0].Mode);
            Assert.AreEqual("Period1", jitterSheet.Rows[0].Period);
            Assert.AreEqual("1.0", jitterSheet.Rows[0].AmplitudeInUi);
            Assert.AreEqual("Test comment", jitterSheet.Rows[0].Comment);
        }

        [TestMethod]
        public void JitterSheet_ClearRows()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            jitterSheet.AddRow(new JitterRow { JitterSet = "JitterSet1" });
            jitterSheet.AddRow(new JitterRow { JitterSet = "JitterSet2" });
            Assert.AreEqual(2, jitterSheet.Rows.Count);

            // Act
            jitterSheet.Rows.Clear();

            // Assert
            Assert.AreEqual(0, jitterSheet.Rows.Count);
        }

        [TestMethod]
        public void JitterSheet_Write_WithEmptyRows_DoesNotCreateFile()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"JitterSheet_Empty_{Guid.NewGuid()}.igx");

            try
            {
                // Act
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
                jitterSheet.Write(tempFileName);

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
        public void JitterSheet_Write_WithRows_DoesNotThrow()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            var jitterRow = new JitterRow
            {
                JitterSet = "JitterSet1",
                PinOrGroup = "Pin1",
                Mode = "Mode1",
                Period = "1.0",
                AmplitudeInUi = "0.5",
                SampleType = "Continuous",
                TargetResolution = "0.01",
                TargetPatternReps = "1000",
                BitsPerPattern = "64",
                MinPatternSkips = "0",
                BitPeriod = "1.0",
                Comment = "Test jitter"
            };
            jitterSheet.AddRow(jitterRow);
            string tempFileName = Path.Combine(Path.GetTempPath(), $"JitterSheet_Data_{Guid.NewGuid()}.igx");

            try
            {
                // Act & Assert - should not throw
                jitterSheet.Write(tempFileName);
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
        public void JitterSheet_Write_WithVersion_DoesNotThrow()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            var jitterRow = new JitterRow
            {
                JitterSet = "JitterSet1",
                PinOrGroup = "Pin1",
                Mode = "Mode1"
            };
            jitterSheet.AddRow(jitterRow);
            string tempFileName = Path.Combine(Path.GetTempPath(), $"JitterSheet_Version_{Guid.NewGuid()}.igx");

            try
            {
                // Act & Assert - should not throw
                jitterSheet.Write(tempFileName, "2.0");
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
        public void JitterSheet_Write_MultipleRows_DoesNotThrow()
        {
            // Arrange
            var jitterSheet = new JitterSheet("Jitter");
            var rows = new List<JitterRow>
            {
                new() { JitterSet = "Set1", PinOrGroup = "Pin1" },
                new() { JitterSet = "Set2", PinOrGroup = "Pin2" },
                new() { JitterSet = "Set3", PinOrGroup = "Pin3" }
            };
            jitterSheet.AddRows(rows);
            string tempFileName = Path.Combine(Path.GetTempPath(), $"JitterSheet_Multiple_{Guid.NewGuid()}.igx");

            try
            {
                // Act & Assert - should not throw
                jitterSheet.Write(tempFileName);
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
