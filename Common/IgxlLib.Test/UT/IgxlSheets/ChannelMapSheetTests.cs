using System.Collections.Generic;

using IgxlLib.Const;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class ChannelMapSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void ChannelMapSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "ChannelMap";

            // Act
            var channelMapSheet = new ChannelMapSheet(sheetName);

            // Assert
            Assert.IsNotNull(channelMapSheet);
            Assert.AreEqual(sheetName, channelMapSheet.Name);
            Assert.AreEqual("DTChanMap", channelMapSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.ChannelMap, channelMapSheet.IgxlSheetName);
        }

        [TestMethod]
        public void ChannelMapSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var channelMapSheet = new ChannelMapSheet("ChannelMap");

            // Assert
            Assert.AreEqual("DTChanMap", channelMapSheet.SheetType);
        }

        [TestMethod]
        public void ChannelMapSheet_Errors_InitializedEmpty()
        {
            // Arrange & Act
            var channelMapSheet = new ChannelMapSheet("ChannelMap");

            // Assert
            Assert.IsNotNull(channelMapSheet.GetErrors());
            Assert.AreEqual(0, channelMapSheet.GetErrors().Count);
        }

        [TestMethod]
        public void ChannelMapSheet_Name_CanBeSet()
        {
            // Arrange
            var channelMapSheet = new ChannelMapSheet("ChannelMap")
            {
                // Act
                Name = "NewChannelMapName"
            };

            // Assert
            Assert.AreEqual("NewChannelMapName", channelMapSheet.Name);
        }

        [TestMethod]
        public void ChannelMapSheet_GetIgxlSheetsVersion()
        {
            // Arrange & Act
            Dictionary<string, Dictionary<string, IGDataXML.IGXLSheetsVersion.SheetInfo>> versionDict = ChannelMapSheet.GetIgxlSheetsVersion();

            // Assert
            Assert.IsNotNull(versionDict);
            Assert.IsTrue(versionDict.Count > 0);
        }
    }
}
