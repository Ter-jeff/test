using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets.MultiSheet.MultiTimeSet
{
    [TestClass]
    public class TimeRow2P3ConverterTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void TimeRow2P3Converter_Constructor()
        {
            // Arrange & Act
            var converter = new TimeRow2P3Converter();

            // Assert
            Assert.IsNotNull(converter);
        }

        [TestMethod]
        public void TimeRow2P3Converter_NeedCompensate_FalseWhenSufficientColumns()
        {
            // Arrange
            var converter = new TimeRow2P3Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17"];

            // Act
            bool result = converter.NeedCompensate(dataArr);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void TimeRow2P3Converter_NeedCompensate_TrueWhenInsufficientColumns()
        {
            // Arrange
            var converter = new TimeRow2P3Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"];

            // Act
            bool result = converter.NeedCompensate(dataArr);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void TimeRow2P3Converter_ConvertTimeRow_With17Columns()
        {
            // Arrange
            var converter = new TimeRow2P3Converter();
            string[] dataArr =
            [
                "Name",
                "SetName",
                "Reserved",
                "PinGrp",
                "100ns",
                "10ns",
                "DataSrc",
                "DataFmt",
                "DriveOn",
                "DriveData",
                "DriveReturn",
                "DriveOff",
                "CompareMode",
                "CompareOpen",
                "CompareClose",
                "CompareRefOffset",
                "EdgeMode"
            ];

            // Act
            TimingRow row = converter.ConvertTimeRow(dataArr);

            // Assert
            Assert.IsNotNull(row);
            Assert.AreEqual("PinGrp", row.PinGrpName);
            Assert.AreEqual("100ns", row.PinGrpClockPeriod);
            Assert.AreEqual("10ns", row.PinGrpSetup);
            Assert.AreEqual("DataSrc", row.DataSrc);
            Assert.AreEqual("DataFmt", row.DataFmt);
            Assert.AreEqual("CompareRefOffset", row.CompareRefOffset);
            Assert.AreEqual("EdgeMode", row.EdgeMode);
            Assert.AreEqual("", row.Comment);
        }

        [TestMethod]
        public void TimeRow2P3Converter_ConvertTimeRow_With18Columns()
        {
            // Arrange
            var converter = new TimeRow2P3Converter();
            string[] dataArr =
            [
                "Name",
                "SetName",
                "Reserved",
                "PinGrp",
                "100ns",
                "10ns",
                "DataSrc",
                "DataFmt",
                "DriveOn",
                "DriveData",
                "DriveReturn",
                "DriveOff",
                "CompareMode",
                "CompareOpen",
                "CompareClose",
                "CompareRefOffset",
                "EdgeMode",
                "Comment"
            ];

            // Act
            TimingRow row = converter.ConvertTimeRow(dataArr);

            // Assert
            Assert.IsNotNull(row);
            Assert.AreEqual("CompareRefOffset", row.CompareRefOffset);
            Assert.AreEqual("EdgeMode", row.EdgeMode);
            Assert.AreEqual("Comment", row.Comment);
        }

        [TestMethod]
        public void TimeRow2P3Converter_ConvertTimeRow_WithInsufficientData()
        {
            // Arrange
            var converter = new TimeRow2P3Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"];

            // Act
            TimingRow row = converter.ConvertTimeRow(dataArr);

            // Assert
            Assert.IsNotNull(row);
            // The array should be compensated to 17 elements
            Assert.AreEqual("4", row.PinGrpName);
        }

        [TestMethod]
        public void TimeRow2P3Converter_InheritsTimeRow1P4Converter()
        {
            // Arrange & Act
            var converter = new TimeRow2P3Converter();

            // Assert
            Assert.IsInstanceOfType(converter, typeof(TimeRow1P4Converter));
        }

        [TestMethod]
        public void TimeRow2P3Converter_CompareRefOffset_IsSet()
        {
            // Arrange
            var converter = new TimeRow2P3Converter();
            string[] dataArr =
            [
                "Name",
                "SetName",
                "Reserved",
                "PinGrp",
                "100ns",
                "10ns",
                "DataSrc",
                "DataFmt",
                "DriveOn",
                "DriveData",
                "DriveReturn",
                "DriveOff",
                "CompareMode",
                "CompareOpen",
                "CompareClose",
                "MyRefOffset",
                "EdgeMode",
                "Comment"
            ];

            // Act
            TimingRow row = converter.ConvertTimeRow(dataArr);

            // Assert
            Assert.AreEqual("MyRefOffset", row.CompareRefOffset);
        }

        [TestMethod]
        public void TimeRow2P3Converter_MustHaveColumnCnt_Is17()
        {
            // Arrange
            var converter = new TimeRow2P3Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17"];

            // Act
            bool result = converter.NeedCompensate(dataArr);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void TimeRow2P3Converter_MustHaveColumnCnt_Is16ThrowsTrue()
        {
            // Arrange
            var converter = new TimeRow2P3Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"];

            // Act
            bool result = converter.NeedCompensate(dataArr);

            // Assert
            Assert.IsTrue(result);
        }
    }
}
