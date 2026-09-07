using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class PortMapSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void PortMapSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "PortMap";

            // Act
            var portMapSheet = new PortMapSheet(sheetName);

            // Assert
            Assert.IsNotNull(portMapSheet);
            Assert.AreEqual(sheetName, portMapSheet.Name);
            Assert.AreEqual("DTPortMapSheet", portMapSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.PortMap, portMapSheet.IgxlSheetName);
        }

        [TestMethod]
        public void PortMapSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var portMapSheet = new PortMapSheet("PortMap");

            // Assert
            Assert.AreEqual("DTPortMapSheet", portMapSheet.SheetType);
        }

        [TestMethod]
        public void PortMapSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var portMapSheet = new PortMapSheet("PortMap");

            // Assert
            Assert.IsNotNull(portMapSheet.GetErrors());
            Assert.AreEqual(0, portMapSheet.GetErrors().Count);
        }

        [TestMethod]
        public void PortMapSheet_Name_CanBeSet()
        {
            // Arrange
            var portMapSheet = new PortMapSheet("PortMap")
            {
                // Act
                Name = "NewPortMapName"
            };

            // Assert
            Assert.AreEqual("NewPortMapName", portMapSheet.Name);
        }

        [TestMethod]
        public void PortMapSheet_GetIgxlSheetsVersion()
        {
            // Arrange & Act
            Dictionary<string, Dictionary<string, IGDataXML.IGXLSheetsVersion.SheetInfo>> versionDict = PortMapSheet.GetIgxlSheetsVersion();

            // Assert
            Assert.IsNotNull(versionDict);
            Assert.IsTrue(versionDict.Count > 0);
        }
    }
}
