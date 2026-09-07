using Automation.Reader.ConfigFile.TimingFileCategoryMapping;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class TimingFileCategoryMappingConfigTests
    {
        private ExcelPackage _package = null!;
        private ExcelWorksheet _sheet = null!;

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("TestMapping");
            _sheet.Cells[1, 1].Value = "CategoryA";
        }

        [TestCleanup]
        public void Cleanup() => _package?.Dispose();

        [TestMethod]
        public void LoadConfig_WhenSheetIsEmpty_ReturnsEmptyConfig()
        {
            // Act
            var config = TimingFileCategoryMappingConfig.LoasConfig(_sheet);

            // Assert
            Assert.AreEqual("", config.GetCategoryMapping("ANY"));
        }

        [TestMethod]
        public void LoadConfig_WithValidData_MapsTimingFileToCategory()
        {
            // Arrange
            // Row 1: Headers (Categories)
            _sheet.Cells[1, 1].Value = "CategoryA";
            _sheet.Cells[1, 2].Value = "CategoryB";
            // Row 2: Data (Timing Files)
            _sheet.Cells[2, 1].Value = "File1";
            _sheet.Cells[2, 2].Value = "File2";

            // Act
            var config = TimingFileCategoryMappingConfig.LoasConfig(_sheet);

            // Assert
            Assert.AreEqual("CATEGORYA", config.GetCategoryMapping("FILE1"));
            // Case-insensitive check
            Assert.AreEqual("CATEGORYB", config.GetCategoryMapping("file2"));
        }

        [TestMethod]
        [DataRow("FILE_X", "CAT_X")]
        [DataRow("file_x", "CAT_X")]
        public void GetCategoryMapping_IsCaseInsensitive(string input, string expected)
        {
            // Arrange
            _sheet.Cells[1, 1].Value = "CAT_X";
            _sheet.Cells[2, 1].Value = "FILE_X";
            var config = TimingFileCategoryMappingConfig.LoasConfig(_sheet);

            // Act
            string result = config.GetCategoryMapping(input);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void GetCategoryMapping_WhenNotFound_ReturnsEmptyString()
        {
            // Arrange
            var config = TimingFileCategoryMappingConfig.LoasConfig(_sheet);

            // Act
            string result = config.GetCategoryMapping("NonExistent");

            // Assert
            Assert.AreEqual("", result);
        }
    }
}
