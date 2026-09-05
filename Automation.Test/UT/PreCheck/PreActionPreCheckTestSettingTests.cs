using Automation.PreCheck.PreChecks.PreAction;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OfficeOpenXml;

using TestPlanLib.DataStruct;

namespace Automation.Test.UT.PreCheck
{
    [TestClass]
    public class PreActionPreCheckTestSettingTests
    {
        private PreActionPreCheckTestSetting _target = null!;

        [TestInitialize]
        public void Setup()
        {
            ExcelWorkbook workbook = new ExcelPackage().Workbook;
            _target = new PreActionPreCheckTestSetting(workbook, "VoltageTable_Test");
        }

        [TestMethod]
        public void HasSpecialUnit_DefaultsToFalse()
        {
            // Act
            bool result = _target.HasSpecialUnit();

            // Assert
            Assert.IsFalse(result);
        }

        [DataTestMethod]
        [DataRow("Category_LV", CategoryValueType.LV)]
        [DataRow("Category_HV", CategoryValueType.HV)]
        [DataRow("Category_NV", CategoryValueType.NV)]
        [DataRow("Other", CategoryValueType.NV)]
        public void GetValueType_MapsStringToCategoryValueType(string value, CategoryValueType expected)
        {
            // Act
            CategoryValueType result = _target.GetValueType(value);

            // Assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void AddFatalErrorColumn_MarksColumnAsFatal()
        {
            // Act
            _target.AddFatalErrorColumn(5);
            bool result = _target.IsFatalErrorColumn(5);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void AddFatalErrorColumn_CalledTwice_DoesNotDuplicate()
        {
            // Act
            _target.AddFatalErrorColumn(5);
            _target.AddFatalErrorColumn(5);

            // Assert
            Assert.AreEqual(1, _target._fatalErrorColumns.Count);
        }

        [TestMethod]
        public void IsFatalErrorColumn_ColumnNotAdded_ReturnsFalse()
        {
            // Act
            bool result = _target.IsFatalErrorColumn(7);

            // Assert
            Assert.IsFalse(result);
        }

        #region CompareTwoCategoryValue

        [TestMethod]
        public void CompareTwoCategoryValue_HvLessThanOther_ReturnsFalseWithLessThanConnector()
        {
            // Act
            bool flag = _target.CompareTwoCategoryValue("1.0", "2.0", CategoryValueType.HV, CategoryValueType.NV, out string connect);

            // Assert
            Assert.IsFalse(flag);
            Assert.AreEqual("<", connect);
        }

        [TestMethod]
        public void CompareTwoCategoryValue_HvGreaterOrEqualThanOther_ReturnsTrue()
        {
            // Act
            bool flag = _target.CompareTwoCategoryValue("2.0", "1.0", CategoryValueType.HV, CategoryValueType.NV, out string connect);

            // Assert
            Assert.IsTrue(flag);
            Assert.AreEqual("X", connect);
        }

        [TestMethod]
        public void CompareTwoCategoryValue_NvGreaterThanHv_ReturnsFalseWithGreaterThanConnector()
        {
            // Act
            bool flag = _target.CompareTwoCategoryValue("2.0", "1.0", CategoryValueType.NV, CategoryValueType.HV, out string connect);

            // Assert
            Assert.IsFalse(flag);
            Assert.AreEqual(">", connect);
        }

        [TestMethod]
        public void CompareTwoCategoryValue_NvLessThanLv_ReturnsFalseWithLessThanConnector()
        {
            // Act
            bool flag = _target.CompareTwoCategoryValue("1.0", "2.0", CategoryValueType.NV, CategoryValueType.LV, out string connect);

            // Assert
            Assert.IsFalse(flag);
            Assert.AreEqual("<", connect);
        }

        [TestMethod]
        public void CompareTwoCategoryValue_LvGreaterThanOther_ReturnsFalseWithGreaterThanConnector()
        {
            // Act
            bool flag = _target.CompareTwoCategoryValue("2.0", "1.0", CategoryValueType.LV, CategoryValueType.NV, out string connect);

            // Assert
            Assert.IsFalse(flag);
            Assert.AreEqual(">", connect);
        }

        [TestMethod]
        public void CompareTwoCategoryValue_NonNumericValues_ReturnsTrueWithXConnector()
        {
            // Act
            bool flag = _target.CompareTwoCategoryValue("abc", "def", CategoryValueType.HV, CategoryValueType.NV, out string connect);

            // Assert
            Assert.IsTrue(flag);
            Assert.AreEqual("X", connect);
        }

        [TestMethod]
        public void CompareTwoCategoryValue_NvVersusNv_AlwaysReturnsTrue()
        {
            // Act - NV vs NV falls through every branch since valueType2 is neither HV nor LV
            bool flag = _target.CompareTwoCategoryValue("5.0", "1.0", CategoryValueType.NV, CategoryValueType.NV, out string connect);

            // Assert
            Assert.IsTrue(flag);
            Assert.AreEqual("X", connect);
        }

        [TestMethod]
        public void CompareTwoCategoryValue_ValuesWithinOrder_ReturnsTrue()
        {
            // Act - HV(2.0) >= NV(1.0): correctly ordered
            bool flag = _target.CompareTwoCategoryValue("2.0", "1.0", CategoryValueType.HV, CategoryValueType.NV, out string connect);

            // Assert
            Assert.IsTrue(flag);
            Assert.AreEqual("X", connect);
        }

        #endregion

        #region GetMaxColumnIndex / GetMaxRowIndex / IsBlankColumn

        [TestMethod]
        public void IsBlankColumn_AllCellsEmpty_ReturnsTrue()
        {
            // Arrange
            ExcelWorkbook workbook = new ExcelPackage().Workbook;
            ExcelWorksheet sheet = workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Header";
            sheet.Cells[2, 1].Value = "PinA";
            var target = new PreActionPreCheckTestSetting(workbook, "Sheet1")
            {
                StartRow = 1,
                StartColumn = 1,
                ExcelWorksheet = sheet
            };

            // Act
            bool result = target.IsBlankColumn(2);

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void IsBlankColumn_HasNonEmptyCell_ReturnsFalse()
        {
            // Arrange
            ExcelWorkbook workbook = new ExcelPackage().Workbook;
            ExcelWorksheet sheet = workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Header";
            sheet.Cells[2, 2].Value = "Value";
            var target = new PreActionPreCheckTestSetting(workbook, "Sheet1")
            {
                StartRow = 1,
                StartColumn = 1,
                ExcelWorksheet = sheet
            };

            // Act
            bool result = target.IsBlankColumn(2);

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void GetMaxColumnIndex_TrailingBlankColumns_ReturnsLastNonBlankColumn()
        {
            // Arrange
            ExcelWorkbook workbook = new ExcelPackage().Workbook;
            ExcelWorksheet sheet = workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Header";
            sheet.Cells[2, 2].Value = "Value";
            sheet.Cells[1, 4].Value = "";
            var target = new PreActionPreCheckTestSetting(workbook, "Sheet1")
            {
                StartRow = 1,
                StartColumn = 1,
                ExcelWorksheet = sheet
            };

            // Act
            int result = target.GetMaxColumnIndex();

            // Assert
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void GetMaxRowIndex_BlankRowInMiddle_ReturnsRowBeforeBlank()
        {
            // Arrange
            ExcelWorkbook workbook = new ExcelPackage().Workbook;
            ExcelWorksheet sheet = workbook.Worksheets.Add("Sheet1");
            sheet.Cells[1, 1].Value = "Header";
            sheet.Cells[2, 1].Value = "PinA";
            sheet.Cells[3, 1].Value = "";
            sheet.Cells[4, 1].Value = "PinB";
            var target = new PreActionPreCheckTestSetting(workbook, "Sheet1")
            {
                StartRow = 1,
                StartColumn = 1,
                ExcelWorksheet = sheet
            };

            // Act
            int result = target.GetMaxRowIndex();

            // Assert
            Assert.AreEqual(2, result);
        }

        #endregion
    }
}
