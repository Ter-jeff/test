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
    public class WaveDefinitionSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void WaveDefinitionSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "WaveDefinition";

            // Act
            var waveDefSheet = new WaveDefinitionSheet(sheetName);

            // Assert
            Assert.IsNotNull(waveDefSheet);
            Assert.AreEqual(sheetName, waveDefSheet.Name);
            Assert.AreEqual("DTWaveDefinitionSheet", waveDefSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.WaveDefinition, waveDefSheet.IgxlSheetName);
        }

        [TestMethod]
        public void WaveDefinitionSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");

            // Assert
            Assert.AreEqual("DTWaveDefinitionSheet", waveDefSheet.SheetType);
        }

        [TestMethod]
        public void WaveDefinitionSheet_IgxlSheetName_IsCorrect()
        {
            // Arrange & Act
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");

            // Assert
            Assert.AreEqual(IgxlSheetNames.WaveDefinition, waveDefSheet.IgxlSheetName);
        }

        [TestMethod]
        public void WaveDefinitionSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");

            // Assert
            Assert.IsNotNull(waveDefSheet.GetErrors());
            Assert.AreEqual(0, waveDefSheet.GetErrors().Count);
        }

        [TestMethod]
        public void WaveDefinitionSheet_Name_CanBeSet()
        {
            // Arrange
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition")
            {
                // Act
                Name = "NewWaveDefinitionName"
            };

            // Assert
            Assert.AreEqual("NewWaveDefinitionName", waveDefSheet.Name);
        }

        [TestMethod]
        public void WaveDefinitionSheet_EmptyRows()
        {
            // Arrange & Act
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");

            // Assert
            Assert.AreEqual(0, waveDefSheet.Rows.Count);
        }

        [TestMethod]
        public void WaveDefinitionSheet_AddRow()
        {
            // Arrange
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");
            var waveDefRow = new WaveDefRow
            {
                WaveDefName = "Wave1",
                WaveDefType = "TypeA",
                RepeatCount = "5"
            };

            // Act
            waveDefSheet.AddRow(waveDefRow);

            // Assert
            Assert.AreEqual(1, waveDefSheet.Rows.Count);
            Assert.AreEqual("Wave1", waveDefSheet.Rows[0].WaveDefName);
        }

        [TestMethod]
        public void WaveDefinitionSheet_AddMultipleRows()
        {
            // Arrange
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");
            var rows = new List<WaveDefRow>
            {
                new() { WaveDefName = "Wave1", WaveDefType = "Type1" },
                new() { WaveDefName = "Wave2", WaveDefType = "Type2" },
                new() { WaveDefName = "Wave3", WaveDefType = "Type3" }
            };

            // Act
            waveDefSheet.AddRows(rows);

            // Assert
            Assert.AreEqual(3, waveDefSheet.Rows.Count);
        }

        [TestMethod]
        public void WaveDefinitionSheet_RowProperties()
        {
            // Arrange
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");
            var waveDefRow = new WaveDefRow
            {
                WaveDefName = "Wave1",
                WaveDefType = "TypeA",
                WaveDefComponent = "Component1",
                RepeatCount = "5",
                RelativePeriod = "10ns",
                RelativeAmplitude = "0.5V",
                RelativeOffset = "0.1V",
                PrimitiveParameters = "Param1",
                Comment = "Test wave"
            };

            // Act
            waveDefSheet.AddRow(waveDefRow);

            // Assert
            Assert.AreEqual("Wave1", waveDefSheet.Rows[0].WaveDefName);
            Assert.AreEqual("TypeA", waveDefSheet.Rows[0].WaveDefType);
            Assert.AreEqual("Component1", waveDefSheet.Rows[0].WaveDefComponent);
            Assert.AreEqual("5", waveDefSheet.Rows[0].RepeatCount);
            Assert.AreEqual("10ns", waveDefSheet.Rows[0].RelativePeriod);
            Assert.AreEqual("0.5V", waveDefSheet.Rows[0].RelativeAmplitude);
            Assert.AreEqual("0.1V", waveDefSheet.Rows[0].RelativeOffset);
            Assert.AreEqual("Param1", waveDefSheet.Rows[0].PrimitiveParameters);
            Assert.AreEqual("Test wave", waveDefSheet.Rows[0].Comment);
        }

        [TestMethod]
        public void WaveDefinitionSheet_Write_CreatesFile()
        {
            // Arrange
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");
            var waveDefRow = new WaveDefRow
            {
                WaveDefName = "Wave1",
                WaveDefType = "TypeA",
                WaveDefComponent = "Component1",
                RepeatCount = "5",
                RelativePeriod = "10ns",
                RelativeAmplitude = "0.5V",
                RelativeOffset = "0.1V",
                PrimitiveParameters = "Param1"
            };
            waveDefSheet.AddRow(waveDefRow);

            string tempFilePath = Path.Combine(Path.GetTempPath(), "test_wavedefinition.txt");
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }

            try
            {
                // Act
                waveDefSheet.Write(tempFilePath);

                // Assert
                Assert.IsTrue(File.Exists(tempFilePath));
                string content = File.ReadAllText(tempFilePath);
                Assert.IsTrue(content.Contains("Wave1"));
                Assert.IsTrue(content.Contains("WaveDefinitionSheet"));
            }
            finally
            {
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
        }

        [TestMethod]
        public void WaveDefinitionSheet_Write_WithMultipleRows()
        {
            // Arrange
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");
            for (int i = 1; i <= 3; i++)
            {
                var waveDefRow = new WaveDefRow
                {
                    WaveDefName = $"Wave{i}",
                    WaveDefType = $"Type{i}",
                    RepeatCount = i.ToString()
                };
                waveDefSheet.AddRow(waveDefRow);
            }

            string tempFilePath = Path.Combine(Path.GetTempPath(), "test_wavedefinition_multi.txt");
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }

            try
            {
                // Act
                waveDefSheet.Write(tempFilePath);

                // Assert
                Assert.IsTrue(File.Exists(tempFilePath));
                string content = File.ReadAllText(tempFilePath);
                Assert.IsTrue(content.Contains("Wave1"));
                Assert.IsTrue(content.Contains("Wave2"));
                Assert.IsTrue(content.Contains("Wave3"));
            }
            finally
            {
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
        }

        [TestMethod]
        public void WaveDefinitionSheet_GetIgxlSheetsVersion()
        {
            // Arrange & Act
            Dictionary<string, Dictionary<string, IGDataXML.IGXLSheetsVersion.SheetInfo>> versionDict = WaveDefinitionSheet.GetIgxlSheetsVersion();

            // Assert
            Assert.IsNotNull(versionDict);
            Assert.IsTrue(versionDict.Count > 0);
        }

        [TestMethod]
        public void WaveDefinitionSheet_GetIgxlSheetsVersion_ContainsWaveDefinition()
        {
            // Arrange & Act
            Dictionary<string, Dictionary<string, IGDataXML.IGXLSheetsVersion.SheetInfo>> versionDict = WaveDefinitionSheet.GetIgxlSheetsVersion();

            // Assert
            Assert.IsNotNull(versionDict);
            Assert.IsTrue(versionDict.Count > 0);
            Assert.IsTrue(versionDict.Keys.Count > 0);
        }

        [TestMethod]
        public void WaveDefinitionSheet_Constructor_WithDifferentNames()
        {
            // Arrange
            string[] sheetNames = ["Wave1", "Wave2", "CustomWaveName"];

            // Act & Assert
            foreach (string name in sheetNames)
            {
                var waveDefSheet = new WaveDefinitionSheet(name);
                Assert.AreEqual(name, waveDefSheet.Name);
                Assert.AreEqual("DTWaveDefinitionSheet", waveDefSheet.SheetType);
            }
        }

        [TestMethod]
        public void WaveDefinitionSheet_AddRow_WithNullProperties()
        {
            // Arrange
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");
            var waveDefRow = new WaveDefRow
            {
                WaveDefName = "Wave1",
                WaveDefType = "TypeA"
                // Other properties left as null/default
            };

            // Act
            waveDefSheet.AddRow(waveDefRow);

            // Assert
            Assert.AreEqual(1, waveDefSheet.Rows.Count);
            Assert.AreEqual("Wave1", waveDefSheet.Rows[0].WaveDefName);
        }

        [TestMethod]
        public void WaveDefinitionSheet_AddRows_WithEmptyList()
        {
            // Arrange
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");
            var emptyRows = new List<WaveDefRow>();

            // Act
            waveDefSheet.AddRows(emptyRows);

            // Assert
            Assert.AreEqual(0, waveDefSheet.Rows.Count);
        }

        [TestMethod]
        public void WaveDefinitionSheet_Rows_CollectionType()
        {
            // Arrange & Act
            var waveDefSheet = new WaveDefinitionSheet("WaveDefinition");

            // Assert
            Assert.IsInstanceOfType(waveDefSheet.Rows, typeof(List<WaveDefRow>));
        }

        [TestMethod]
        public void WaveDefinitionSheet_MultipleInstances_AreIndependent()
        {
            // Arrange
            var sheet1 = new WaveDefinitionSheet("Sheet1");
            var sheet2 = new WaveDefinitionSheet("Sheet2");

            sheet1.AddRow(new WaveDefRow { WaveDefName = "Wave1" });
            sheet2.AddRow(new WaveDefRow { WaveDefName = "Wave2" });

            // Act & Assert
            Assert.AreEqual(1, sheet1.Rows.Count);
            Assert.AreEqual(1, sheet2.Rows.Count);
            Assert.AreEqual("Sheet1", sheet1.Name);
            Assert.AreEqual("Sheet2", sheet2.Name);
        }
    }
}
