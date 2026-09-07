using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class MainFlowTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void MainFlow_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "MainFlow";

            // Act
            var mainFlow = new MainFlow(sheetName);

            // Assert
            Assert.IsNotNull(mainFlow);
            Assert.AreEqual(sheetName, mainFlow.Name);
            Assert.AreEqual("DTFlowtableSheet", mainFlow.SheetType);
            Assert.AreEqual(IgxlSheetNames.FlowTable, mainFlow.IgxlSheetName);
        }

        [TestMethod]
        public void MainFlow_AddRow()
        {
            // Arrange
            var mainFlow = new MainFlow("MainFlow");
            var flowRow = new FlowRow { Parameter = "FlowName" };

            // Act
            mainFlow.AddRow(flowRow);

            // Assert
            Assert.AreEqual(1, mainFlow.Rows.Count);
        }

        [TestMethod]
        public void MainFlow_AddRows()
        {
            // Arrange
            var mainFlow = new MainFlow("MainFlow");
            var rows = new List<FlowRow>
            {
                new() { Parameter = "FlowName1"},
                new() { Parameter = "FlowName2"},
                new() { Parameter = "FlowName3"}
            };

            // Act
            mainFlow.AddRows(rows);

            // Assert
            Assert.AreEqual(3, mainFlow.Rows.Count);
        }

        [TestMethod]
        public void MainFlow_RemoveRow()
        {
            // Arrange
            var mainFlow = new MainFlow("MainFlow");
            var row1 = new FlowRow { Parameter = "FlowName1" };
            var row2 = new FlowRow { Parameter = "FlowName2" };
            mainFlow.AddRow(row1);
            mainFlow.AddRow(row2);

            // Act
            mainFlow.RemoveRow(row1);

            // Assert
            Assert.AreEqual(1, mainFlow.Rows.Count);
            Assert.AreEqual("FlowName2", mainFlow.Rows[0].Parameter);
        }

        [TestMethod]
        public void MainFlow_InsertRow()
        {
            // Arrange
            var mainFlow = new MainFlow("MainFlow");
            var row1 = new FlowRow { Parameter = "FlowName1" };
            var row2 = new FlowRow { Parameter = "FlowName2" };
            var rowToInsert = new FlowRow { Parameter = "FlowName3" };
            mainFlow.AddRow(row1);
            mainFlow.AddRow(row2);

            // Act
            int index = mainFlow.InsertRow(1, rowToInsert);

            // Assert
            Assert.AreEqual(2, index);
            Assert.AreEqual(3, mainFlow.Rows.Count);
            Assert.AreEqual("FlowName3", mainFlow.Rows[1].Parameter);
        }

        [TestMethod]
        public void MainFlow_SheetType_IsCorrect()
        {
            // Arrange & Act
            var mainFlow = new MainFlow("MainFlow");

            // Assert
            Assert.AreEqual("DTFlowtableSheet", mainFlow.SheetType);
        }

        [TestMethod]
        public void MainFlow_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var mainFlow = new MainFlow("MainFlow");

            // Assert
            Assert.IsNotNull(mainFlow.GetErrors());
            Assert.AreEqual(0, mainFlow.GetErrors().Count);
        }

        [TestMethod]
        public void MainFlow_Name_CanBeSet()
        {
            // Arrange
            var mainFlow = new MainFlow("MainFlow")
            {
                // Act
                Name = "NewMainFlowName"
            };

            // Assert
            Assert.AreEqual("NewMainFlowName", mainFlow.Name);
        }

        [TestMethod]
        public void MainFlow_GetIgxlSheetsVersion()
        {
            // Arrange & Act
            Dictionary<string, Dictionary<string, IGDataXML.IGXLSheetsVersion.SheetInfo>> versionDict = MainFlow.GetIgxlSheetsVersion();

            // Assert
            Assert.IsNotNull(versionDict);
            Assert.IsTrue(versionDict.Count > 0);
        }
    }
}
