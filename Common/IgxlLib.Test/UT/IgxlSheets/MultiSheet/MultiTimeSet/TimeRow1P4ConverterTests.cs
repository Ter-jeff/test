using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets.MultiSheet.MultiTimeSet
{
    [TestClass]
    public class TimeRow1P4ConverterTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void TimeRow1P4Converter_Constructor()
        {
            // Arrange & Act
            var converter = new TimeRow1P4Converter();

            // Assert
            Assert.IsNotNull(converter);
        }

        [TestMethod]
        public void TimeRow1P4Converter_NeedCompensate_FalseWhenSufficientColumns()
        {
            // Arrange
            var converter = new TimeRow1P4Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"];

            // Act
            bool result = converter.NeedCompensate(dataArr);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void TimeRow1P4Converter_NeedCompensate_TrueWhenInsufficientColumns()
        {
            // Arrange
            var converter = new TimeRow1P4Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"];

            // Act
            bool result = converter.NeedCompensate(dataArr);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void TimeRow1P4Converter_ConvertTimeRow_WithValidData()
        {
            // Arrange
            var converter = new TimeRow1P4Converter();
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
                "EdgeMode",
                "Comment"
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
            Assert.AreEqual("DriveOn", row.DriveOn);
            Assert.AreEqual("DriveData", row.DriveData);
            Assert.AreEqual("DriveReturn", row.DriveReturn);
            Assert.AreEqual("DriveOff", row.DriveOff);
            Assert.AreEqual("CompareMode", row.CompareMode);
            Assert.AreEqual("CompareOpen", row.CompareOpen);
            Assert.AreEqual("CompareClose", row.CompareClose);
            Assert.AreEqual("EdgeMode", row.EdgeMode);
            Assert.AreEqual("Comment", row.Comment);
        }

        [TestMethod]
        public void TimeRow1P4Converter_ConvertTimeRow_WithInsufficientData()
        {
            // Arrange
            var converter = new TimeRow1P4Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"];

            // Act
            TimingRow row = converter.ConvertTimeRow(dataArr);

            // Assert
            Assert.IsNotNull(row);
            // The array should be compensated to 16 elements
            Assert.AreEqual("4", row.PinGrpName);
        }

        [TestMethod]
        public void TimeRow1P4Converter_ConvertTimeRow_WithMinimumRequiredColumns()
        {
            // Arrange
            var converter = new TimeRow1P4Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"];

            // Act
            TimingRow row = converter.ConvertTimeRow(dataArr);

            // Assert
            Assert.IsNotNull(row);
            Assert.AreEqual("", row.Comment);
        }

        [TestMethod]
        public void TimeRow1P4Converter_ConvertTimeRow_WithMoreThanMinimumColumns()
        {
            // Arrange
            var converter = new TimeRow1P4Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "Comment", "Extra1", "Extra2", "Extra3"];

            // Act
            TimingRow row = converter.ConvertTimeRow(dataArr);

            // Assert
            Assert.IsNotNull(row);
            Assert.AreEqual("Comment", row.Comment);
        }

        [TestMethod]
        public void TimeRow1P4Converter_ConvertTimeRow_EmptyComment()
        {
            // Arrange
            var converter = new TimeRow1P4Converter();
            string[] dataArr = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"];

            // Act
            TimingRow row = converter.ConvertTimeRow(dataArr);

            // Assert
            Assert.AreEqual("", row.Comment);
        }
    }
}
