using Automation.GenerateIgxl.EFuse.InputChecker;
using Automation.Static;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestPlanLib.Efuse.Input;

namespace Automation.Test.UT.Efuse
{

    [TestClass]
    public class EfuseArraySizeCheckerTests
    {
        private EfuseArraySizeChecker _checker = null!;
        private EfuseArraySizeSheet _sheet = null!;

        [TestInitialize]
        public void Setup()
        {
            _checker = new EfuseArraySizeChecker();
            _sheet = new EfuseArraySizeSheet("SheetName")
            {
                Rows = [],
                SheetName = "TestSheet"
            };
            // Reset device to a default non-RF state
            LocalSpecs.Options.Device = EnumDevice.AP;
            _sheet.ClearErrors();
        }

        [TestMethod]
        public void CheckOfBitColumn_WhenEmpty_AddsError()
        {
            // Arrange: Line 41-44 branch
            _sheet.Rows.Add(new EfuseArraySizeRow { OfBits = "", RowNum = 1 });

            // Act
            _checker.WorkFlow(_sheet);

            // Assert
            Assert.AreEqual(1, _sheet.GetErrors().Count, "Should report error for empty column");
        }

        [TestMethod]
        public void CheckOfBitColumn_WhenNotNumeric_AddsError()
        {
            // Arrange: Line 48-51 branch
            _sheet.Rows.Add(new EfuseArraySizeRow { OfBits = "ABC", RowNum = 2 });

            // Act
            _checker.WorkFlow(_sheet);

            // Assert
            Assert.AreEqual(1, _sheet.GetErrors().Count, "Should report error for non-numeric input");
        }

        [TestMethod]
        public void CheckOfBitCountIsDivideBy16_WhenRFAndNotDivisible_AddsError()
        {
            // Arrange: Line 17 (RF Device) and Line 28 (Mod 16 check)
            LocalSpecs.Options.Device = EnumDevice.RF;
            // 10 % 16 != 0
            _sheet.Rows.Add(new EfuseArraySizeRow { OfBits = "10", RowNum = 3 });

            // Act
            _checker.WorkFlow(_sheet);

            // Assert
            Assert.AreEqual(1, _sheet.GetErrors().Count, "RF devices must have bit counts divisible by 16");
        }

        [TestMethod]
        public void CheckOfBitCountIsDivideBy16_WhenRFAndValid_NoErrors()
        {
            // Arrange: Valid RF case
            LocalSpecs.Options.Device = EnumDevice.RF;
            // 32 % 16 == 0
            _sheet.Rows.Add(new EfuseArraySizeRow { OfBits = "32", RowNum = 4 });

            // Act
            _checker.WorkFlow(_sheet);

            // Assert
            Assert.AreEqual(0, _sheet.GetErrors().Count);
        }
    }
}
