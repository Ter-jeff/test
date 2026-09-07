using Automation.Singleton;
using Automation.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

namespace Automation.Test.UT.Common
{
    [TestClass]
    public class ExceptionListSingletonTests
    {
        private ExcelPackage _package = null!;
        private ExcelWorksheet _sheet = null!;

        [TestInitialize]
        public void Setup()
        {
            // Reset the Singleton instance
            ExceptionListSingleton._instance = null;

            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("InstanceExceptionList");

            // Mock the global workbook reference
            EpWorkbook.TestPlanWorkbook = _package.Workbook;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package?.Dispose();
            EpWorkbook.TestPlanWorkbook = null;
        }

        [TestMethod]
        public void GetInstance_ShouldAlwaysReturnSameObject()
        {
            var instance1 = ExceptionListSingleton.Instance();
            var instance2 = ExceptionListSingleton.Instance();

            Assert.AreSame(instance1, instance2);
        }

        [TestMethod]
        public void Initialize_WithValidData_LoadsExceptionList()
        {
            // Arrange: Header at 1,1; Data starts 2 rows below (Row 3)
            _sheet.Cells[1, 1].Value = "ExceptionList";
            // TestName
            _sheet.Cells[3, 1].Value = "Test_A";
            // DcCategory
            _sheet.Cells[3, 2].Value = "DC_Cat_1";

            var singleton = ExceptionListSingleton.Instance();

            // Act
            singleton.Initialize();
            string category = singleton.GetDcCategoryByInstance("Test_A");

            // Assert
            Assert.AreEqual("DC_Cat_1", category);
        }

        [TestMethod]
        public void GetDcCategoryByInstance_WhenNotFound_ReturnsEmptyString()
        {
            // Act
            string result = ExceptionListSingleton.Instance().GetDcCategoryByInstance("UnknownTest");

            // Assert
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void FindCellByValue_ShouldReturnCorrectCoordinates()
        {
            // Arrange
            _sheet.Cells[5, 10].Value = "TargetValue";
            int row = 0, col = 0;

            // Act
            ExceptionListSingleton.FindCellByValue(ref row, ref col, _sheet, "targetvalue");

            // Assert
            Assert.AreEqual(5, row);
            Assert.AreEqual(10, col);
        }
    }
}
